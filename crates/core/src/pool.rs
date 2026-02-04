//! LibreOffice process pool for parallel document conversion.
//!
//! This module manages a pool of LibreOffice instances for converting
//! Office documents to PDF. Each instance runs in its own process with
//! a separate user profile to enable true parallel execution.

use crate::config::PoolConfig;
use crate::error::{ConversionError, Result};
use async_channel::{bounded, Receiver, Sender};
use async_process::Command;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicU32, AtomicUsize, Ordering};
use std::sync::Arc;
use std::time::Instant;
use tempfile::TempDir;
use tokio::sync::{Mutex, Semaphore};
use tokio::time::{timeout, Duration};
use tracing::{debug, error, info, warn};
use uuid::Uuid;

/// A job to be processed by the pool.
struct ConversionJob {
    /// Input file path.
    input_path: PathBuf,
    /// Response channel for the result.
    response_tx: tokio::sync::oneshot::Sender<Result<PathBuf>>,
}

/// A single LibreOffice instance in the pool.
struct LibreOfficeInstance {
    /// Instance ID for logging.
    id: usize,
    /// Unique user profile directory (required for parallel execution).
    profile_dir: TempDir,
    /// Number of documents processed by this instance.
    docs_processed: AtomicU32,
    /// Whether this instance is currently processing.
    is_busy: AtomicBool,
}

impl std::fmt::Debug for LibreOfficeInstance {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("LibreOfficeInstance")
            .field("id", &self.id)
            .field("profile_dir", &self.profile_dir.path())
            .field("docs_processed", &self.docs_processed.load(Ordering::SeqCst))
            .field("is_busy", &self.is_busy.load(Ordering::SeqCst))
            .finish()
    }
}

impl LibreOfficeInstance {
    /// Create a new LibreOffice instance with a unique profile.
    fn new(id: usize) -> Result<Self> {
        let profile_dir = TempDir::with_prefix(&format!("lo-profile-{}-", id))
            .map_err(|e| ConversionError::ProcessStartFailed(e))?;

        debug!(
            "Created LibreOffice instance {} with profile at {:?}",
            id,
            profile_dir.path()
        );

        Ok(Self {
            id,
            profile_dir,
            docs_processed: AtomicU32::new(0),
            is_busy: AtomicBool::new(false),
        })
    }

    /// Get the profile directory path.
    fn profile_path(&self) -> &Path {
        self.profile_dir.path()
    }

    /// Increment the document counter and return the new value.
    fn increment_docs(&self) -> u32 {
        self.docs_processed.fetch_add(1, Ordering::SeqCst) + 1
    }

    /// Get the number of documents processed.
    fn docs_processed(&self) -> u32 {
        self.docs_processed.load(Ordering::SeqCst)
    }

    /// Check if this instance needs recycling.
    fn needs_recycling(&self, max_docs: u32) -> bool {
        self.docs_processed() >= max_docs
    }

    /// Mark as busy.
    fn set_busy(&self, busy: bool) {
        self.is_busy.store(busy, Ordering::SeqCst);
    }
}

/// Pool of LibreOffice instances for parallel document conversion.
#[derive(Debug)]
pub struct LibreOfficePool {
    /// Pool configuration.
    config: PoolConfig,
    /// Path to soffice binary.
    soffice_path: PathBuf,
    /// Instances in the pool.
    instances: Vec<Arc<Mutex<LibreOfficeInstance>>>,
    /// Semaphore to limit concurrent conversions.
    semaphore: Arc<Semaphore>,
    /// Whether the pool is shut down.
    is_shutdown: AtomicBool,
    /// Temporary directory for output PDFs.
    output_temp_dir: TempDir,
    /// Total documents processed.
    total_processed: AtomicUsize,
}

impl LibreOfficePool {
    /// Create a new LibreOffice pool.
    pub async fn new(config: PoolConfig) -> Result<Self> {
        config.validate()?;

        // Find soffice binary
        let soffice_path = Self::find_soffice(&config)?;
        info!("Found LibreOffice at: {:?}", soffice_path);

        // Create instances
        let mut instances = Vec::with_capacity(config.pool_size);
        for i in 0..config.pool_size {
            let instance = LibreOfficeInstance::new(i)?;
            instances.push(Arc::new(Mutex::new(instance)));
        }

        // Create temp directory for PDFs
        let output_temp_dir = TempDir::with_prefix("office-to-png-pdfs-")
            .map_err(|e| ConversionError::ProcessStartFailed(e))?;

        info!(
            "LibreOffice pool initialized with {} instances",
            config.pool_size
        );

        let pool_size = config.pool_size;
        Ok(Self {
            config,
            soffice_path,
            instances,
            semaphore: Arc::new(Semaphore::new(pool_size)),
            is_shutdown: AtomicBool::new(false),
            output_temp_dir,
            total_processed: AtomicUsize::new(0),
        })
    }

    /// Find the soffice binary.
    fn find_soffice(config: &PoolConfig) -> Result<PathBuf> {
        // Check if explicit path is provided
        if let Some(ref path) = config.soffice_path {
            if path.exists() {
                return Ok(path.clone());
            }
            return Err(ConversionError::LibreOfficeNotFound);
        }

        // Search common locations
        let candidates = [
            // macOS
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            // Linux
            "/usr/bin/soffice",
            "/usr/lib/libreoffice/program/soffice",
            "/opt/libreoffice/program/soffice",
            // Snap (Ubuntu)
            "/snap/bin/libreoffice.soffice",
        ];

        for candidate in candidates {
            let path = PathBuf::from(candidate);
            if path.exists() {
                return Ok(path);
            }
        }

        // Try PATH
        which::which("soffice")
            .or_else(|_| which::which("libreoffice"))
            .map_err(|_| ConversionError::LibreOfficeNotFound)
    }

    /// Convert a document to PDF.
    pub async fn convert_to_pdf(&self, input_path: &Path) -> Result<PathBuf> {
        if self.is_shutdown.load(Ordering::SeqCst) {
            return Err(ConversionError::PoolShutdown);
        }

        // Validate input
        if !input_path.exists() {
            return Err(ConversionError::InputNotFound(input_path.to_path_buf()));
        }

        // Validate extension
        let ext = input_path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("");
        if !crate::is_supported_extension(ext) {
            return Err(ConversionError::UnsupportedFormat {
                extension: ext.to_string(),
            });
        }

        // Acquire semaphore permit
        let _permit = self
            .semaphore
            .acquire()
            .await
            .map_err(|_| ConversionError::PoolShutdown)?;

        // Find an available instance
        let instance = self.get_available_instance().await?;

        // Run conversion
        let result = self
            .run_conversion(&instance, input_path)
            .await;

        // Release instance
        {
            let inst = instance.lock().await;
            inst.set_busy(false);
        }

        result
    }

    /// Get an available instance from the pool.
    async fn get_available_instance(&self) -> Result<Arc<Mutex<LibreOfficeInstance>>> {
        for instance in &self.instances {
            let inst = instance.lock().await;
            if !inst.is_busy.load(Ordering::SeqCst) {
                inst.set_busy(true);
                drop(inst);
                return Ok(Arc::clone(instance));
            }
        }

        // This shouldn't happen due to semaphore, but handle gracefully
        Err(ConversionError::PoolExhausted {
            pool_size: self.config.pool_size,
        })
    }

    /// Run the actual LibreOffice conversion.
    async fn run_conversion(
        &self,
        instance: &Arc<Mutex<LibreOfficeInstance>>,
        input_path: &Path,
    ) -> Result<PathBuf> {
        let start = Instant::now();
        
        // Get instance info
        let (instance_id, profile_path) = {
            let inst = instance.lock().await;
            (inst.id, inst.profile_path().to_path_buf())
        };

        debug!(
            "Instance {} converting {:?}",
            instance_id,
            input_path.file_name()
        );

        // Create unique output directory for this conversion
        let output_dir = self.output_temp_dir.path().join(Uuid::new_v4().to_string());
        std::fs::create_dir_all(&output_dir).map_err(|e| ConversionError::OutputDirError {
            path: output_dir.clone(),
            message: e.to_string(),
        })?;

        // Build command
        let mut cmd = Command::new(&self.soffice_path);
        cmd.args([
            "--headless",
            "--invisible",
            "--nologo",
            "--nofirststartwizard",
            "--norestore",
        ]);

        // Set unique user profile (critical for parallel execution!)
        cmd.arg(format!(
            "-env:UserInstallation=file://{}",
            profile_path.display()
        ));

        // Conversion arguments
        cmd.args([
            "--convert-to",
            "pdf:writer_pdf_Export", // Use writer export for best quality
            "--outdir",
        ]);
        cmd.arg(&output_dir);
        cmd.arg(input_path);

        // Run with timeout
        let output = timeout(self.config.conversion_timeout, cmd.output())
            .await
            .map_err(|_| ConversionError::Timeout {
                path: input_path.to_path_buf(),
                timeout_secs: self.config.conversion_timeout.as_secs(),
            })?
            .map_err(|e| ConversionError::ProcessStartFailed(e))?;

        // Check exit status
        if !output.status.success() {
            let stderr = String::from_utf8_lossy(&output.stderr);
            error!(
                "LibreOffice conversion failed for {:?}: {}",
                input_path, stderr
            );
            return Err(ConversionError::ConversionFailed {
                path: input_path.to_path_buf(),
                message: stderr.to_string(),
            });
        }

        // Find the output PDF
        let input_stem = input_path
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("output");
        let pdf_path = output_dir.join(format!("{}.pdf", input_stem));

        if !pdf_path.exists() {
            // LibreOffice might have created a file with slightly different name
            // Try to find any PDF in the output directory
            let pdf = std::fs::read_dir(&output_dir)
                .ok()
                .and_then(|entries| {
                    entries
                        .filter_map(|e| e.ok())
                        .find(|e| {
                            e.path()
                                .extension()
                                .map(|ext| ext == "pdf")
                                .unwrap_or(false)
                        })
                        .map(|e| e.path())
                });

            if let Some(found_pdf) = pdf {
                debug!(
                    "Instance {} converted {:?} in {:?}",
                    instance_id,
                    input_path.file_name(),
                    start.elapsed()
                );

                // Update counters
                {
                    let inst = instance.lock().await;
                    inst.increment_docs();
                }
                self.total_processed.fetch_add(1, Ordering::SeqCst);

                return Ok(found_pdf);
            }

            return Err(ConversionError::ConversionFailed {
                path: input_path.to_path_buf(),
                message: "PDF output file not found".to_string(),
            });
        }

        debug!(
            "Instance {} converted {:?} in {:?}",
            instance_id,
            input_path.file_name(),
            start.elapsed()
        );

        // Update counters
        {
            let inst = instance.lock().await;
            inst.increment_docs();
        }
        self.total_processed.fetch_add(1, Ordering::SeqCst);

        Ok(pdf_path)
    }

    /// Convert multiple documents in parallel.
    pub async fn convert_batch(&self, input_paths: Vec<PathBuf>) -> Vec<Result<PathBuf>> {
        use futures::future::join_all;

        let futures: Vec<_> = input_paths
            .iter()
            .map(|path| self.convert_to_pdf(path))
            .collect();

        join_all(futures).await
    }

    /// Get pool health information.
    pub async fn health(&self) -> PoolHealth {
        let mut instances_info = Vec::with_capacity(self.instances.len());

        for instance in &self.instances {
            let inst = instance.lock().await;
            instances_info.push(InstanceHealth {
                id: inst.id,
                docs_processed: inst.docs_processed(),
                is_busy: inst.is_busy.load(Ordering::SeqCst),
                needs_recycling: inst.needs_recycling(self.config.max_docs_per_instance),
            });
        }

        PoolHealth {
            pool_size: self.config.pool_size,
            total_processed: self.total_processed.load(Ordering::SeqCst),
            is_shutdown: self.is_shutdown.load(Ordering::SeqCst),
            instances: instances_info,
        }
    }

    /// Shutdown the pool.
    pub async fn shutdown(&self) {
        info!("Shutting down LibreOffice pool");
        self.is_shutdown.store(true, Ordering::SeqCst);
        // Temp directories will be cleaned up when dropped
    }

    /// Get the total number of documents processed.
    pub fn total_processed(&self) -> usize {
        self.total_processed.load(Ordering::SeqCst)
    }
}

/// Health information for the pool.
#[derive(Debug, Clone)]
pub struct PoolHealth {
    /// Total pool size.
    pub pool_size: usize,
    /// Total documents processed.
    pub total_processed: usize,
    /// Whether the pool is shut down.
    pub is_shutdown: bool,
    /// Per-instance health info.
    pub instances: Vec<InstanceHealth>,
}

/// Health information for a single instance.
#[derive(Debug, Clone)]
pub struct InstanceHealth {
    /// Instance ID.
    pub id: usize,
    /// Documents processed by this instance.
    pub docs_processed: u32,
    /// Whether currently busy.
    pub is_busy: bool,
    /// Whether needs recycling.
    pub needs_recycling: bool,
}

#[cfg(test)]
mod tests {
    use super::*;

    // ========== LibreOfficeInstance tests ==========

    #[test]
    fn test_instance_creation() {
        let instance = LibreOfficeInstance::new(0).unwrap();
        assert_eq!(instance.id, 0);
        assert_eq!(instance.docs_processed(), 0);
        assert!(!instance.is_busy.load(Ordering::SeqCst));
        assert!(instance.profile_path().exists());
    }

    #[test]
    fn test_instance_profile_path_is_unique() {
        let instance1 = LibreOfficeInstance::new(0).unwrap();
        let instance2 = LibreOfficeInstance::new(1).unwrap();
        assert_ne!(instance1.profile_path(), instance2.profile_path());
    }

    #[test]
    fn test_instance_increment_docs() {
        let instance = LibreOfficeInstance::new(0).unwrap();
        assert_eq!(instance.docs_processed(), 0);
        
        assert_eq!(instance.increment_docs(), 1);
        assert_eq!(instance.docs_processed(), 1);
        
        assert_eq!(instance.increment_docs(), 2);
        assert_eq!(instance.docs_processed(), 2);
    }

    #[test]
    fn test_instance_needs_recycling() {
        let instance = LibreOfficeInstance::new(0).unwrap();
        
        // Fresh instance doesn't need recycling
        assert!(!instance.needs_recycling(100));
        
        // Simulate processing docs
        for _ in 0..100 {
            instance.increment_docs();
        }
        
        // Now should need recycling
        assert!(instance.needs_recycling(100));
        assert!(instance.needs_recycling(50)); // exceeded 50
        assert!(!instance.needs_recycling(200)); // not reached 200
    }

    #[test]
    fn test_instance_busy_flag() {
        let instance = LibreOfficeInstance::new(0).unwrap();
        
        assert!(!instance.is_busy.load(Ordering::SeqCst));
        
        instance.set_busy(true);
        assert!(instance.is_busy.load(Ordering::SeqCst));
        
        instance.set_busy(false);
        assert!(!instance.is_busy.load(Ordering::SeqCst));
    }

    // ========== PoolHealth tests ==========

    #[test]
    fn test_pool_health_struct() {
        let health = PoolHealth {
            pool_size: 4,
            total_processed: 100,
            is_shutdown: false,
            instances: vec![
                InstanceHealth {
                    id: 0,
                    docs_processed: 25,
                    is_busy: false,
                    needs_recycling: false,
                },
                InstanceHealth {
                    id: 1,
                    docs_processed: 25,
                    is_busy: true,
                    needs_recycling: false,
                },
            ],
        };

        assert_eq!(health.pool_size, 4);
        assert_eq!(health.total_processed, 100);
        assert!(!health.is_shutdown);
        assert_eq!(health.instances.len(), 2);
    }

    #[test]
    fn test_instance_health_clone() {
        let health = InstanceHealth {
            id: 5,
            docs_processed: 42,
            is_busy: true,
            needs_recycling: false,
        };
        
        let cloned = health.clone();
        assert_eq!(cloned.id, 5);
        assert_eq!(cloned.docs_processed, 42);
        assert!(cloned.is_busy);
        assert!(!cloned.needs_recycling);
    }

    // ========== LibreOfficePool tests (may require LO installed) ==========

    #[tokio::test]
    async fn test_pool_creation() {
        let config = PoolConfig::with_pool_size(2);
        // This test will fail if LibreOffice is not installed
        // That's expected - it's an integration test
        let result = LibreOfficePool::new(config).await;
        
        // Just check that we get a reasonable error if LO isn't installed
        match result {
            Ok(_) => (), // LibreOffice is installed
            Err(ConversionError::LibreOfficeNotFound) => (),
            Err(e) => panic!("Unexpected error: {:?}", e),
        }
    }

    #[tokio::test]
    async fn test_pool_rejects_zero_size() {
        let mut config = PoolConfig::default();
        config.pool_size = 0;
        
        let result = LibreOfficePool::new(config).await;
        match result {
            Ok(_) => panic!("Expected error for zero pool size"),
            Err(ConversionError::InvalidConfig(_)) => (),
            Err(e) => panic!("Expected InvalidConfig, got {:?}", e),
        }
    }

    #[tokio::test]
    async fn test_pool_with_nonexistent_soffice_path() {
        let mut config = PoolConfig::default();
        config.soffice_path = Some(PathBuf::from("/nonexistent/path/to/soffice"));
        
        let result = LibreOfficePool::new(config).await;
        match result {
            Ok(_) => panic!("Expected error for nonexistent soffice path"),
            Err(ConversionError::LibreOfficeNotFound) => (),
            Err(e) => panic!("Expected LibreOfficeNotFound, got {:?}", e),
        }
    }

    // ========== find_soffice tests ==========

    #[test]
    fn test_find_soffice_with_explicit_nonexistent_path() {
        let mut config = PoolConfig::default();
        config.soffice_path = Some(PathBuf::from("/nonexistent/soffice"));
        
        let result = LibreOfficePool::find_soffice(&config);
        assert!(matches!(result, Err(ConversionError::LibreOfficeNotFound)));
    }

    #[test]
    fn test_find_soffice_with_explicit_valid_path() {
        // Use a path that definitely exists (this binary itself during tests)
        let current_exe = std::env::current_exe().unwrap();
        
        let mut config = PoolConfig::default();
        config.soffice_path = Some(current_exe.clone());
        
        let result = LibreOfficePool::find_soffice(&config);
        assert!(result.is_ok());
        assert_eq!(result.unwrap(), current_exe);
    }
}

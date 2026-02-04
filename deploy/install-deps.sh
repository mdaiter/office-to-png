#!/bin/bash
# EC2 Deployment Script for office-to-png
# Supports: Amazon Linux 2023, Ubuntu 22.04/24.04
#
# Usage: ./install-deps.sh [--with-pdfium]

set -euo pipefail

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

log_info() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

log_warn() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# Detect OS
detect_os() {
    if [ -f /etc/os-release ]; then
        . /etc/os-release
        OS=$ID
        VERSION=$VERSION_ID
    else
        log_error "Cannot detect OS"
        exit 1
    fi
    log_info "Detected OS: $OS $VERSION"
}

# Install dependencies on Amazon Linux 2023
install_amazon_linux() {
    log_info "Installing dependencies for Amazon Linux 2023..."
    
    # Update system
    sudo dnf update -y
    
    # Install development tools
    sudo dnf groupinstall -y "Development Tools"
    sudo dnf install -y cmake gcc-c++ pkgconfig
    
    # Install LibreOffice
    log_info "Installing LibreOffice..."
    sudo dnf install -y libreoffice-core libreoffice-writer libreoffice-calc libreoffice-impress
    
    # Install fonts
    log_info "Installing fonts..."
    sudo dnf install -y \
        liberation-fonts \
        dejavu-fonts-common \
        google-noto-fonts-common \
        google-noto-sans-fonts \
        fontconfig
    
    # Install Rust
    install_rust
    
    # Install Python
    log_info "Installing Python..."
    sudo dnf install -y python3 python3-pip python3-devel
    
    # Install pdfium (optional)
    if [ "${INSTALL_PDFIUM:-false}" = "true" ]; then
        install_pdfium_linux
    fi
}

# Install dependencies on Ubuntu
install_ubuntu() {
    log_info "Installing dependencies for Ubuntu..."
    
    # Update system
    sudo apt-get update
    sudo apt-get upgrade -y
    
    # Install development tools
    sudo apt-get install -y build-essential cmake pkg-config curl
    
    # Install LibreOffice
    log_info "Installing LibreOffice..."
    sudo apt-get install -y \
        libreoffice-core \
        libreoffice-writer \
        libreoffice-calc \
        libreoffice-impress \
        --no-install-recommends
    
    # Install fonts
    log_info "Installing fonts..."
    sudo apt-get install -y \
        fonts-liberation \
        fonts-dejavu \
        fonts-noto \
        fontconfig
    
    # Install MS core fonts (requires accepting EULA)
    log_info "Installing Microsoft core fonts..."
    echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | \
        sudo debconf-set-selections
    sudo apt-get install -y ttf-mscorefonts-installer || log_warn "MS fonts not installed"
    
    # Rebuild font cache
    sudo fc-cache -fv
    
    # Install Rust
    install_rust
    
    # Install Python
    log_info "Installing Python..."
    sudo apt-get install -y python3 python3-pip python3-venv python3-dev
    
    # Install pdfium (optional)
    if [ "${INSTALL_PDFIUM:-false}" = "true" ]; then
        install_pdfium_linux
    fi
}

# Install Rust
install_rust() {
    if command -v rustc &> /dev/null; then
        log_info "Rust already installed: $(rustc --version)"
        return
    fi
    
    log_info "Installing Rust..."
    curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y
    source "$HOME/.cargo/env"
    
    # Install wasm target for browser builds
    rustup target add wasm32-unknown-unknown
    
    # Install useful tools
    cargo install wasm-pack maturin
    
    log_info "Rust installed: $(rustc --version)"
}

# Install pdfium library
install_pdfium_linux() {
    log_info "Installing pdfium..."
    
    # Download pre-built pdfium
    PDFIUM_VERSION="6392"
    ARCH=$(uname -m)
    
    case "$ARCH" in
        x86_64)
            PDFIUM_ARCH="linux-x64"
            ;;
        aarch64)
            PDFIUM_ARCH="linux-arm64"
            ;;
        *)
            log_error "Unsupported architecture: $ARCH"
            return 1
            ;;
    esac
    
    PDFIUM_URL="https://github.com/nickel-chromium/nickel-chromium/releases/download/pdfium-${PDFIUM_VERSION}/pdfium-${PDFIUM_ARCH}.tgz"
    
    # Download and extract
    curl -L "$PDFIUM_URL" -o /tmp/pdfium.tgz
    sudo mkdir -p /usr/local/lib/pdfium
    sudo tar -xzf /tmp/pdfium.tgz -C /usr/local/lib/pdfium
    
    # Set up library path
    echo "/usr/local/lib/pdfium/lib" | sudo tee /etc/ld.so.conf.d/pdfium.conf
    sudo ldconfig
    
    # Export environment variable
    export PDFIUM_LIB_DIR="/usr/local/lib/pdfium/lib"
    echo "export PDFIUM_LIB_DIR=/usr/local/lib/pdfium/lib" >> ~/.bashrc
    
    log_info "pdfium installed to /usr/local/lib/pdfium"
}

# Configure system for high throughput
configure_system() {
    log_info "Configuring system for high throughput..."
    
    # Increase file descriptor limits
    echo "* soft nofile 65535" | sudo tee -a /etc/security/limits.conf
    echo "* hard nofile 65535" | sudo tee -a /etc/security/limits.conf
    
    # Set up tmpfs for temp files (optional, for ephemeral instances)
    if [ "${USE_TMPFS:-false}" = "true" ]; then
        log_info "Setting up tmpfs for /tmp..."
        echo "tmpfs /tmp tmpfs defaults,noatime,mode=1777,size=4G 0 0" | sudo tee -a /etc/fstab
    fi
    
    # Configure LibreOffice for headless operation
    mkdir -p ~/.config/libreoffice/4/user
    cat > ~/.config/libreoffice/4/user/registrymodifications.xcu << 'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry"
           xmlns:xs="http://www.w3.org/2001/XMLSchema"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <item oor:path="/org.openoffice.Office.Common/Misc">
    <prop oor:name="UseSystemPrintDialog" oor:op="fuse">
      <value>false</value>
    </prop>
  </item>
</oor:items>
EOF
}

# Verify installation
verify_installation() {
    log_info "Verifying installation..."
    
    local errors=0
    
    # Check LibreOffice
    if command -v soffice &> /dev/null; then
        log_info "✓ LibreOffice: $(soffice --version 2>/dev/null | head -1)"
    else
        log_error "✗ LibreOffice not found"
        ((errors++))
    fi
    
    # Check Rust
    if command -v rustc &> /dev/null; then
        log_info "✓ Rust: $(rustc --version)"
    else
        log_error "✗ Rust not found"
        ((errors++))
    fi
    
    # Check Python
    if command -v python3 &> /dev/null; then
        log_info "✓ Python: $(python3 --version)"
    else
        log_error "✗ Python not found"
        ((errors++))
    fi
    
    # Check fonts
    if fc-list | grep -q "Liberation"; then
        log_info "✓ Liberation fonts installed"
    else
        log_warn "○ Liberation fonts not found"
    fi
    
    # Check pdfium
    if [ -f "/usr/local/lib/pdfium/lib/libpdfium.so" ]; then
        log_info "✓ pdfium library found"
    else
        log_warn "○ pdfium not installed (will use bundled version)"
    fi
    
    if [ $errors -gt 0 ]; then
        log_error "Installation completed with $errors errors"
        return 1
    fi
    
    log_info "Installation completed successfully!"
}

# Print usage
print_usage() {
    echo "Usage: $0 [OPTIONS]"
    echo ""
    echo "Options:"
    echo "  --with-pdfium    Install system pdfium library"
    echo "  --with-tmpfs     Configure tmpfs for /tmp"
    echo "  --help           Show this help message"
    echo ""
    echo "Example:"
    echo "  $0 --with-pdfium"
}

# Main
main() {
    INSTALL_PDFIUM=false
    USE_TMPFS=false
    
    while [[ $# -gt 0 ]]; do
        case $1 in
            --with-pdfium)
                INSTALL_PDFIUM=true
                shift
                ;;
            --with-tmpfs)
                USE_TMPFS=true
                shift
                ;;
            --help)
                print_usage
                exit 0
                ;;
            *)
                log_error "Unknown option: $1"
                print_usage
                exit 1
                ;;
        esac
    done
    
    export INSTALL_PDFIUM
    export USE_TMPFS
    
    detect_os
    
    case "$OS" in
        amzn|amazon)
            install_amazon_linux
            ;;
        ubuntu|debian)
            install_ubuntu
            ;;
        *)
            log_error "Unsupported OS: $OS"
            exit 1
            ;;
    esac
    
    configure_system
    verify_installation
    
    echo ""
    log_info "Next steps:"
    echo "  1. Log out and back in (or run: source ~/.bashrc)"
    echo "  2. Clone your office-to-png repository"
    echo "  3. Build with: cargo build --release"
    echo "  4. Build Python wheel: cd crates/python && maturin build --release"
}

main "$@"

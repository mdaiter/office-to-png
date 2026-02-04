# Multi-stage build for office-to-png-server
# 
# Build: docker build -f deploy/Dockerfile -t office-to-png-server .
# Run:   docker run -p 8080:8080 -e S3_BUCKET=my-bucket office-to-png-server

# ==============================================================================
# Stage 1: Build the Rust binary
# ==============================================================================
FROM rust:1.75-bookworm AS builder

# Install build dependencies
RUN apt-get update && apt-get install -y \
    cmake \
    clang \
    pkg-config \
    libssl-dev \
    && rm -rf /var/lib/apt/lists/*

# Download and install pdfium
# Using the pdfium-binaries release (Google's PDF engine)
RUN mkdir -p /opt/pdfium && \
    curl -L https://github.com/aspect-build/aspect-build-pdfium/releases/download/pdfium-6406/pdfium-linux-x64.tar.gz \
    | tar xz -C /opt/pdfium && \
    cp /opt/pdfium/lib/libpdfium.so /usr/local/lib/ && \
    ldconfig

# Create app directory
WORKDIR /app

# Copy workspace files
COPY Cargo.toml Cargo.lock ./
COPY crates/core/Cargo.toml crates/core/Cargo.toml
COPY crates/server/Cargo.toml crates/server/Cargo.toml
COPY crates/python/Cargo.toml crates/python/Cargo.toml
COPY crates/wasm/Cargo.toml crates/wasm/Cargo.toml
COPY tests/fixtures/generator/Cargo.toml tests/fixtures/generator/Cargo.toml

# Create dummy source files for dependency caching
RUN mkdir -p crates/core/src crates/server/src crates/python/src crates/wasm/src tests/fixtures/generator/src && \
    echo 'pub fn dummy() {}' > crates/core/src/lib.rs && \
    echo 'fn main() {}' > crates/server/src/main.rs && \
    touch crates/server/src/lib.rs && \
    echo 'pub fn dummy() {}' > crates/python/src/lib.rs && \
    echo 'pub fn dummy() {}' > crates/wasm/src/lib.rs && \
    echo 'fn main() {}' > tests/fixtures/generator/src/main.rs

# Build dependencies only (this layer will be cached)
RUN cargo build --release -p office-to-png-server 2>/dev/null || true

# Copy actual source code
COPY crates/core crates/core
COPY crates/server crates/server

# Build the actual binary
RUN cargo build --release -p office-to-png-server

# ==============================================================================
# Stage 2: Runtime image
# ==============================================================================
FROM debian:bookworm-slim

# Install runtime dependencies
RUN apt-get update && apt-get install -y \
    # LibreOffice (headless, minimal)
    libreoffice-writer-nogui \
    libreoffice-calc-nogui \
    libreoffice-impress-nogui \
    # Font support
    fonts-liberation \
    fonts-dejavu-core \
    fontconfig \
    # SSL and other libs
    ca-certificates \
    libssl3 \
    # For debugging (optional, can be removed in production)
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy pdfium library
COPY --from=builder /usr/local/lib/libpdfium.so /usr/local/lib/
RUN ldconfig

# Copy the binary
COPY --from=builder /app/target/release/office-to-png-server /usr/local/bin/

# Create non-root user
RUN useradd -m -u 1000 appuser
USER appuser

# Environment variables
ENV LD_LIBRARY_PATH=/usr/local/lib
ENV RUST_LOG=info,office_to_png_server=debug,office_to_png_core=debug
ENV PORT=8080
ENV HOST=0.0.0.0
ENV POOL_SIZE=2
ENV DPI=150

# Expose port
EXPOSE 8080

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:8080/health || exit 1

# Run the server
CMD ["office-to-png-server"]

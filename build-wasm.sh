#!/bin/bash
set -e

echo "Building WASM version of GridView..."

# Install wasm-bindgen-cli if not already installed
if ! command -v wasm-bindgen &> /dev/null; then
    echo "wasm-bindgen not found. Installing wasm-bindgen-cli..."
    cargo install wasm-bindgen-cli
fi

# Add wasm32 target if not already added
rustup target add wasm32-unknown-unknown

# Build for WASM
echo "Building..."
cargo build --release --target wasm32-unknown-unknown

# Generate wasm-bindgen bindings
echo "Generating bindings..."
wasm-bindgen --out-dir ./dist --target web ./target/wasm32-unknown-unknown/release/csv-app.wasm

# Copy index.html to dist
echo "Copying index.html..."
mkdir -p dist
cp index.html dist/

echo "Build complete! Output is in ./dist/"
echo "To serve locally, run: python3 -m http.server 8080 --directory dist"
echo "Then open http://localhost:8080 in your browser"

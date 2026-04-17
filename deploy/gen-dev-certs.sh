#!/bin/bash
# Generates self-signed dev certs for local testing ONLY.
# For production, replace certs/server.crt and certs/server.key with
# certificates issued by your campus internal CA.

set -e
CERTS_DIR="$(dirname "$0")/../certs"
mkdir -p "$CERTS_DIR"

# Check openssl is available
if ! command -v openssl &> /dev/null; then
  echo "Error: openssl not found. Install it and re-run."
  exit 1
fi

echo "Generating development self-signed certificate..."
openssl req -x509 -nodes -days 365 \
  -newkey rsa:2048 \
  -keyout "$CERTS_DIR/server.key" \
  -out "$CERTS_DIR/server.crt" \
  -subj "/C=US/ST=YourState/L=YourCity/O=YourCollege/CN=localhost" \
  -addext "subjectAltName=DNS:localhost,IP:127.0.0.1"

echo ""
echo "Dev certs written to $CERTS_DIR/"
echo ""
echo "IMPORTANT: These are development-only certificates."
echo "For production, replace with certificates from your campus internal CA."
echo "Trust this cert in Windows: double-click server.crt → Install → Local Machine → Trusted Root"

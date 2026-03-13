#!/usr/bin/env bash
set -e

echo "==> Installing LibreOffice..."
apt-get update -qq
apt-get install -y libreoffice

echo "==> Installing Python dependencies..."
pip install -r requirements.txt

echo "==> Build complete!"

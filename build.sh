#!/bin/bash
echo "==> Installing LibreOffice..."
apt-get update && apt-get install -y libreoffice

echo "==> Installing Python dependencies..."
pip install -r requirements.txt

echo "==> Done! Run: gunicorn app:app"

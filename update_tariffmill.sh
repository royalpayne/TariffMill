#!/bin/bash
# TariffMill Update Script for Ubuntu
# Simple one-command update: ./update_tariffmill.sh

set -e

INSTALL_DIR="$HOME/Dev/Tariffmill"
VENV_DIR="$INSTALL_DIR/.venv"

echo "========================================"
echo "  TariffMill Update"
echo "========================================"

# Check if install directory exists
if [ ! -d "$INSTALL_DIR" ]; then
    echo "TariffMill not found at $INSTALL_DIR"
    echo "Run initial install first."
    exit 1
fi

cd "$INSTALL_DIR"

# Pull latest from GitHub
echo ""
echo "Pulling latest updates from GitHub..."
git pull origin main

# Activate virtual environment
if [ ! -d "$VENV_DIR" ]; then
    echo "Creating virtual environment..."
    python3 -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"

# Install/update the package
echo ""
echo "Installing updates..."
pip install --upgrade pip
pip install -e .

echo ""
echo "========================================"
echo "  Update Complete!"
echo "========================================"
echo ""
echo "Run TariffMill with: tariffmill"
echo "Or use the desktop shortcut."
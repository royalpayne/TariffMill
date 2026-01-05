#!/bin/bash
# TariffMill Install Script for Ubuntu
# First-time installation: ./install_tariffmill.sh

set -e

INSTALL_DIR="$HOME/Dev/Tariffmill"
VENV_DIR="$INSTALL_DIR/.venv"
DESKTOP_FILE="$HOME/Desktop/TariffMill.desktop"

echo "========================================"
echo "  TariffMill Installation"
echo "========================================"

# Check for Python 3
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is required"
    echo "Install with: sudo apt install python3 python3-pip python3-venv"
    exit 1
fi

# Clone or update repo
if [ ! -d "$INSTALL_DIR" ]; then
    echo "Cloning TariffMill repository..."
    mkdir -p "$HOME/Dev"
    git clone https://github.com/ProcessLogicLabs/TariffMill.git "$INSTALL_DIR"
else
    echo "TariffMill already exists, pulling updates..."
    cd "$INSTALL_DIR"
    git pull origin main
fi

cd "$INSTALL_DIR"

# Create virtual environment
echo ""
echo "Setting up virtual environment..."
if [ ! -d "$VENV_DIR" ]; then
    python3 -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"

# Install dependencies
echo ""
echo "Installing dependencies..."
pip install --upgrade pip
pip install -e .

# Create desktop shortcut
echo ""
echo "Creating desktop shortcut..."
cat > "$DESKTOP_FILE" << EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=TariffMill
Comment=Customs Entry Processing Application
Exec=$VENV_DIR/bin/tariffmill
Icon=$INSTALL_DIR/Tariffmill/Resources/tariffmill_icon.png
Terminal=false
Categories=Office;
StartupNotify=true
EOF

chmod +x "$DESKTOP_FILE"

# Mark desktop file as trusted (GNOME)
if command -v gio &> /dev/null; then
    gio set "$DESKTOP_FILE" metadata::trusted true 2>/dev/null || true
fi

echo ""
echo "========================================"
echo "  Installation Complete!"
echo "========================================"
echo ""
echo "You can now run TariffMill:"
echo "  1. Double-click the desktop shortcut"
echo "  2. Or run: $VENV_DIR/bin/tariffmill"
echo ""
echo "To update later, run: $INSTALL_DIR/update_tariffmill.sh"
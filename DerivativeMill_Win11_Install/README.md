# DerivativeMill Windows 11 Installation Package

Complete Windows 11 installation setup for DerivativeMill.

## Quick Start

### Build the Installation Package

**On Windows 11 with Python 3.8+:**

```powershell
cd DerivativeMill_Win11_Install\scripts
powershell -ExecutionPolicy Bypass -File build_windows_installer.ps1
```

Or use the batch script:
```batch
build_windows_installer.bat
```

**Build time**: 2-5 minutes
**Output**: `dist/DerivativeMill_Windows11_Portable.zip`

## Directory Structure

```
DerivativeMill_Win11_Install/
├── README.md                          (This file)
├── scripts/
│   ├── build_windows.spec            (PyInstaller config)
│   ├── build_windows_installer.bat   (Batch build script)
│   ├── build_windows_installer.ps1   (PowerShell build script)
│   └── uninstall_windows.ps1         (Uninstaller)
└── docs/
    ├── WINDOWS_INSTALLATION.md       (User guide)
    ├── WINDOWS_PACKAGE_README.md     (Package info)
    ├── BUILD_WINDOWS_PACKAGE.md      (Build instructions)
    ├── WINDOWS_DEPLOYMENT_SUMMARY.md (Summary)
    └── WINDOWS_PACKAGE_CHECKLIST.md  (Quick reference)
```

## Installation Methods (3 Options)

### Method 1: Portable (No Installation)
```
1. Extract ZIP anywhere
2. Run DerivativeMill.exe
3. No system changes
```

### Method 2: With Desktop Shortcut
```
1. Extract ZIP to location
2. Run INSTALL.bat
3. Desktop shortcut created
```

### Method 3: System-Wide
```
1. Extract to C:\Program Files\DerivativeMill\
2. Run INSTALL.bat
3. Available to all users
```

## System Requirements

- **OS**: Windows 10 (SP1+) or Windows 11
- **RAM**: 4GB minimum, 8GB recommended
- **Disk Space**: 500MB free
- **Display**: 1280x720 minimum

## File Guide

### Scripts Directory

| File | Purpose |
|------|---------|
| `build_windows.spec` | PyInstaller configuration |
| `build_windows_installer.bat` | Batch build script (Windows) |
| `build_windows_installer.ps1` | PowerShell build script (recommended) |
| `uninstall_windows.ps1` | Professional uninstaller |

### Documentation Directory

| File | Audience | Purpose |
|------|----------|---------|
| `WINDOWS_INSTALLATION.md` | End Users | Complete installation guide |
| `WINDOWS_PACKAGE_README.md` | End Users | Package overview and features |
| `BUILD_WINDOWS_PACKAGE.md` | Developers | Build instructions and customization |
| `WINDOWS_DEPLOYMENT_SUMMARY.md` | Administrators | Technical specifications and deployment |
| `WINDOWS_PACKAGE_CHECKLIST.md` | Everyone | Quick reference checklist |

## Building the Package

### PowerShell (Recommended)

```powershell
# Navigate to scripts directory
cd scripts

# Run build script
powershell -ExecutionPolicy Bypass -File build_windows_installer.ps1
```

### Batch Script

```batch
cd scripts
build_windows_installer.bat
```

## What Gets Created

```
dist/
├── DerivativeMill.exe                          (~200 MB)
├── DerivativeMill_Windows11_Portable.zip       (Distribution)
└── DerivativeMill/
    ├── DerivativeMill.exe
    ├── Run_DerivativeMill.bat
    ├── INSTALL.bat
    ├── Input/
    ├── Output/
    ├── ProcessedPDFs/
    └── Resources/derivativemill.db
```

## Documentation

- **Getting Started**: Read `docs/WINDOWS_INSTALLATION.md`
- **Package Overview**: Read `docs/WINDOWS_PACKAGE_README.md`
- **Build Instructions**: Read `docs/BUILD_WINDOWS_PACKAGE.md`
- **Technical Details**: Read `docs/WINDOWS_DEPLOYMENT_SUMMARY.md`
- **Quick Reference**: Read `docs/WINDOWS_PACKAGE_CHECKLIST.md`

## Uninstallation

```powershell
# Professional uninstaller with data backup
powershell -ExecutionPolicy Bypass -File scripts/uninstall_windows.ps1
```

## System Requirements

### For Running

- Windows 10 (SP1+) or Windows 11
- 4GB RAM minimum
- 500MB disk space
- 1280x720 display

### For Building

- Python 3.8+
- PyInstaller 6.17.0
- 2GB disk space for build artifacts

## Version Information

- **Version**: 1.08
- **Created**: December 2024
- **Platform**: Windows 10/11
- **Python**: 3.12 (bundled)
- **Status**: Production Ready

## Support

For issues or questions:

1. Check the appropriate documentation file
2. Review `docs/WINDOWS_INSTALLATION.md` troubleshooting section
3. Check application Log View tab for error messages
4. See `docs/WINDOWS_PACKAGE_CHECKLIST.md` for quick reference

## Features

✓ Automated build scripts (batch and PowerShell)
✓ Multiple installation methods
✓ Professional uninstaller with data backup
✓ Comprehensive documentation
✓ Fully portable (USB compatible)
✓ No system installation required
✓ Complete offline support

---

**Ready to build?** Run `scripts/build_windows_installer.ps1` now!

For detailed instructions, see `docs/WINDOWS_INSTALLATION.md`

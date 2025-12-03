# Building DerivativeMill Portable Windows 11 Package

This document provides step-by-step instructions for building the portable Windows 11 installation package on your Windows 11 PC.

## Prerequisites

- **Windows 11 PC** with administrative access
- **Python 3.8+** installed and added to PATH
- **Project files** (DerivativeMill directory from GitHub)

## Quick Build (5 Minutes)

### Step 1: Prepare the Environment

1. Open **Command Prompt** (cmd.exe) or **PowerShell**
2. Navigate to the project root directory:
   ```batch
   cd C:\path\to\Project_mv
   ```

3. Create a Python virtual environment:
   ```batch
   python -m venv venv
   ```

4. Activate the virtual environment:
   ```batch
   venv\Scripts\activate.bat
   ```

5. Install build dependencies:
   ```batch
   pip install -r requirements.txt
   pip install PyInstaller wheel
   ```

### Step 2: Build the Package

Run the build script from the project root:

```batch
DerivativeMill_Win11_Install\scripts\build_windows_installer.bat
```

The script will:
- ✅ Check Python installation
- ✅ Activate virtual environment
- ✅ Install PyInstaller
- ✅ Build executable (2-5 minutes)
- ✅ Create portable ZIP package
- ✅ Package everything into `dist\DerivativeMill_Windows11_Portable.zip`

### Step 3: Test the Build

After the build completes:

```batch
dist\DerivativeMill\DerivativeMill.exe
```

The application should launch. Test basic functionality and then close it.

## Build Output

After successful build, you'll have:

- **`dist\DerivativeMill\DerivativeMill.exe`** - Standalone executable (~150MB)
- **`dist\DerivativeMill_Windows11_Portable.zip`** - Portable ZIP package (~80MB compressed)
- **`dist\DerivativeMill\INSTALL.bat`** - Installation helper script
- **`dist\DerivativeMill\Run_DerivativeMill.bat`** - Quick launcher script

## Distribution

### Option 1: Portable USB Drive
Simply copy the contents of `dist\DerivativeMill` to a USB drive and run the `.exe` anywhere without installation.

### Option 2: Program Files Installation
1. Create folder: `C:\Program Files\DerivativeMill`
2. Copy `dist\DerivativeMill\*` contents to that folder
3. Run `INSTALL.bat` to create desktop shortcut
4. Run `DerivativeMill.exe`

### Option 3: ZIP Distribution
Share the `dist\DerivativeMill_Windows11_Portable.zip` file:
1. User extracts ZIP anywhere
2. User runs `DerivativeMill.exe`
3. Application creates data folders automatically on first run

## Version Information

- **Current Version**: v0.60.1 (stored in `DerivativeMill/version.py`)
- **To Update Version**:
  - Edit `DerivativeMill/version.py`
  - Change `__version__ = "v0.60.1"` to your new version
  - Run build script again

## Troubleshooting

### "Python not found"
- Install Python 3.8+ from [python.org](https://www.python.org)
- Make sure to check "Add Python to PATH" during installation
- Restart Command Prompt after installing

### "Virtual environment not found"
```batch
python -m venv venv
venv\Scripts\activate.bat
pip install -r requirements.txt
```

### "PyInstaller not found"
```batch
venv\Scripts\activate.bat
pip install PyInstaller wheel
```

### Build fails or takes too long
- Ensure you have 2GB+ free disk space
- Close antivirus software (may slow build)
- Delete `build/` and `dist/` folders and try again

## Build Specifications

- **Format**: PyInstaller one-file bundle
- **Type**: Windowed application (no console)
- **Icon**: `DerivativeMill/Resources/derivativemill.ico`
- **Included Files**:
  - PyQt5 GUI framework
  - pandas (data processing)
  - openpyxl (Excel export)
  - pdfplumber (PDF processing)
  - PIL (Image processing)
  - SQLite database
  - Resource files (icons, images)
  - Documentation (README.md, QUICKSTART.md, SETUP.md)

## File Structure

```
dist/
├── DerivativeMill/
│   ├── DerivativeMill.exe          (Main executable)
│   ├── Run_DerivativeMill.bat      (Quick launcher)
│   ├── INSTALL.bat                 (Setup helper)
│   ├── README.md                   (Project overview)
│   ├── QUICKSTART.md              (Quick start guide)
│   ├── SETUP.md                    (Detailed setup)
│   └── (Data folders created by app on first run)
└── DerivativeMill_Windows11_Portable.zip   (Portable package)
```

## Next Steps

After building:
1. Test the executable thoroughly
2. Create a version tag in Git: `git tag -a v0.60.1 -m "Release v0.60.1"`
3. Push tag to GitHub: `git push origin v0.60.1`
4. Distribute the ZIP file or copy to installation location

## Support

For issues or questions:
- Check BUILD_INSTRUCTIONS.md in DerivativeMill_Win11_Install/
- Review application logs in: `%APPDATA%\DerivativeMill\logs\`
- Ensure all dependencies in requirements.txt are installed

---

**Build Script**: `DerivativeMill_Win11_Install\scripts\build_windows_installer.bat`
**Spec File**: `DerivativeMill_Win11_Install\scripts\build_windows.spec`
**Last Updated**: December 2024

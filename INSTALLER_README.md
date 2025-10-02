# Tower Bolt Tension Data Tool - Installation Guide

## Quick Start

### For End Users (Recommended)

1. **Download the Installer**
   - Download `Install.bat` from the GitHub repository
   - Place it in any folder on your computer

2. **Run the Installer**
   - Double-click `Install.bat`
   - Follow the on-screen prompts
   - The installer will:
     - Download the latest version from GitHub
     - Set up Python environment
     - Install all dependencies
     - Create shortcuts

3. **Launch the Application**
   - Use the desktop shortcut (if created during installation)
   - Or run `Launch.bat` from the installation folder

---

## Installation Methods

### Method 1: Automated Installer (Easiest)

**Requirements:**
- Windows 7 or later
- Python 3.8 or later installed
- Internet connection

**Steps:**
1. Download and run `Install.bat`
2. Choose installation location (or use default)
3. Wait for automatic setup to complete
4. Launch the application

**What it does:**
- ✓ Downloads latest version from GitHub
- ✓ Creates isolated Python environment
- ✓ Installs all required packages
- ✓ Creates launcher scripts
- ✓ Creates desktop shortcut (optional)

---

### Method 2: Update Existing Installation (VBS Script)

If you already have the application installed:

1. Double-click `Install or Update from GitHub.vbs`
2. Confirm the update
3. Wait for download and extraction
4. Files will be updated automatically

**Note:** This updates files only, doesn't reinstall dependencies.

---

## Manual Installation (Advanced Users)

If you prefer manual setup:

```bash
# 1. Clone or download the repository
git clone https://github.com/benchaney53/VestasTensionDataTool.git

# 2. Navigate to the app directory
cd VestasTensionDataTool/app

# 3. Create virtual environment
python -m venv ../venv

# 4. Activate virtual environment
..\venv\Scripts\activate

# 5. Install dependencies
pip install -r requirements.txt

# 6. Run the application
python main.py
```

---

## Troubleshooting

### Python Not Found
- Install Python from https://python.org/downloads/
- During installation, check "Add Python to PATH"
- Restart your computer after installing Python

### Installation Fails
- Make sure you have internet connection
- Check antivirus isn't blocking the download
- Try running `Install.bat` as Administrator

### Application Won't Start
- Make sure Python is installed correctly
- Try reinstalling using the installer
- Check that all dependencies installed successfully

### Permission Errors
- Right-click `Install.bat` and select "Run as Administrator"
- Choose an installation location where you have write permissions

---

## Uninstallation

1. Delete the installation folder
2. Delete the desktop shortcut (if created)
3. No registry changes or system files are modified

---

## Files Included

- `Install.bat` - Main installer launcher (double-click this)
- `installer.ps1` - PowerShell installation script
- `Install or Update from GitHub.vbs` - Update existing installations
- `Run Tower Bolt App.vbs` - Alternative launcher

---

## Support

For issues or questions, please visit:
https://github.com/benchaney53/VestasTensionDataTool

---

## System Requirements

- **OS:** Windows 7 or later
- **Python:** 3.8 or later
- **RAM:** 2GB minimum (4GB recommended)
- **Disk Space:** 500MB for application and dependencies
- **Internet:** Required for installation only


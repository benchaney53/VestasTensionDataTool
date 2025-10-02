# Tower Bolt Tension Data Tool - Installation Guide

## Quick Start

### For End Users (Recommended - Easiest!)

1. **Download the Standalone Installer**
   - Download **`Install-Standalone.bat`** from the GitHub repository
   - Save it **anywhere** on your computer (Desktop, Downloads, etc.)
   - **You only need this ONE file!**

2. **Run the Installer**
   - Double-click `Install-Standalone.bat`
   - Follow the on-screen prompts
   - The installer will automatically:
     - Download the latest version from GitHub
     - Set up Python environment
     - Install all dependencies
     - Create launcher and shortcuts

3. **Launch the Application**
   - Use the desktop shortcut (if created during installation)
   - Or run `Launch.bat` from the installation folder

**Why use the standalone installer?**
- ✅ Only need to download ONE file
- ✅ Works from any location
- ✅ Downloads everything else automatically
- ✅ No manual file management required

---

## Installation Methods

### Method 1: Standalone Installer (Easiest - Recommended!)

**File:** `Install-Standalone.bat`

**Requirements:**
- Windows 7 or later
- Python 3.8 or later installed
- Internet connection

**Steps:**
1. Download **only** `Install-Standalone.bat`
2. Save it anywhere (Desktop, Downloads folder, etc.)
3. Double-click to run
4. Choose installation location
5. Wait for automatic setup
6. Launch the application

**Advantages:**
- ✅ **Only one file to download** - no dependencies
- ✅ **Works from anywhere** - doesn't need other files
- ✅ **Self-contained** - downloads everything automatically
- ✅ **Always gets latest version** from GitHub

---

### Method 2: Full Installer Package

**Files:** `Install.bat` + `installer.ps1` (both files needed)

**Requirements:**
- Windows 7 or later
- Python 3.8 or later installed
- Internet connection

**Steps:**
1. Download both `Install.bat` AND `installer.ps1` to the same folder
2. Run `Install.bat`
3. Choose installation location (or use default)
4. Wait for automatic setup to complete
5. Launch the application

**What it does:**
- ✓ Downloads latest version from GitHub
- ✓ Creates isolated Python environment
- ✓ Installs all required packages
- ✓ Creates launcher scripts
- ✓ Creates desktop shortcut (optional)

**Note:** This method requires both files. Use Method 1 (standalone) if you want a single-file installer.

---

### Method 3: Update Existing Installation (VBS Script)

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

## Installer Files Available

- **`Install-Standalone.bat`** - ⭐ **RECOMMENDED** - Single-file installer (downloads everything)
- `Install.bat` - Main installer launcher (requires installer.ps1)
- `installer.ps1` - PowerShell installation script (used by Install.bat)
- `Install or Update from GitHub.vbs` - Update existing installations
- `Run Tower Bolt App.vbs` - Alternative launcher

**For new users:** Download `Install-Standalone.bat` - it's the easiest option!

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


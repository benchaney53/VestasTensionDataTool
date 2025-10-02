# Tower Bolt Tension Data Tool - Installation Guide

## Quick Start

### For End Users (Recommended - Professional GUI Installer)

1. **Download the Setup Wizard**
   - Download **`Setup.exe.bat`** from the GitHub repository
   - Save it **anywhere** on your computer (Desktop, Downloads, etc.)
   - **You only need this ONE file!**

2. **Run the Professional Setup**
   - Double-click `Setup.exe.bat`
   - Follow the modern Windows wizard interface
   - The installer will automatically:
     - Download the latest version from GitHub
     - Set up Python environment
     - Install all dependencies
     - Create launcher and shortcuts

3. **Launch the Application**
   - Use the desktop shortcut (created automatically)
   - Or run `Launch.bat` from the installation folder

**✨ Why use the GUI Setup Wizard?**
- 🎨 **Professional appearance** - Modern Windows installer interface
- 📋 **Step-by-step guidance** - Clear wizard with progress indicators
- 📝 **License agreement** - Proper EULA acceptance
- 📁 **Location selection** - Choose where to install
- ⚡ **Real-time progress** - See installation status and progress bars
- 🔍 **Automatic Python detection** - Finds and uses the right Python installation
- 🛡️ **Error handling** - Clear error messages and troubleshooting

---

## Installation Methods

### Method 1: GUI Setup Wizard (Recommended - Professional)

**File:** `Setup.exe.bat`

**Requirements:**
- Windows 7 or later
- Python 3.8 or later installed
- Internet connection

**Steps:**
1. Download **only** `Setup.exe.bat`
2. Save it anywhere (Desktop, Downloads folder, etc.)
3. Double-click to run
4. Follow the professional wizard interface
5. Choose installation options
6. Wait for automatic setup
7. Launch the application

**Advantages:**
- 🎨 **Professional GUI** - Modern Windows installer appearance
- 📋 **Guided installation** - Step-by-step wizard with clear instructions
- 📝 **License acceptance** - Proper EULA agreement
- 📁 **Location selection** - Choose custom installation folder
- ⚡ **Progress tracking** - Real-time progress bars and status
- 🔍 **Smart Python detection** - Automatically finds compatible Python
- 🛡️ **Error handling** - Clear error messages and recovery

---

### Method 2: Command-Line Standalone Installer

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

- **`Setup.exe.bat`** - ⭐ **RECOMMENDED** - Professional GUI setup wizard (downloads everything)
- `Setup.exe.ps1` - PowerShell GUI installer script (used by Setup.exe.bat)
- `Install-Standalone.bat` - Command-line single-file installer
- `Install.bat` - Main installer launcher (requires installer.ps1)
- `installer.ps1` - PowerShell installation script (used by Install.bat)
- `Install or Update from GitHub.vbs` - Update existing installations
- `Run Tower Bolt App.vbs` - Alternative launcher

**For new users:** Download `Setup.exe.bat` - it's the most professional and user-friendly option!

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


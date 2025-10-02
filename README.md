# Tower Bolt Tension Data Tool

A professional application for analyzing and reporting tower bolt tension data for Vestas wind turbines.

## Features

- 📊 Tower bolt tension data analysis
- 📈 Visual data representation with charts
- 📄 Automated PDF report generation
- 🔧 Flange calculation utilities
- 🎨 Professional Vestas-branded output
- 💾 Configuration management

## Quick Installation

### For End Users (Easiest Method)

1. **Download** the standalone installer:
   - Download `Install-Standalone.bat` from this repository
   - Save it anywhere on your computer
   
2. **Run** the installer:
   - Double-click `Install-Standalone.bat`
   - Follow the on-screen instructions
   - Everything downloads and installs automatically!
   
3. **Launch** the application:
   - Use the desktop shortcut or run `Launch.bat`

**Note:** The standalone installer downloads everything it needs from GitHub - you only need this one file!

For detailed installation instructions and alternative methods, see [INSTALLER_README.md](INSTALLER_README.md)

## Requirements

- Windows 7 or later
- Python 3.8 or later
- Internet connection (for installation only)

## What's Included

- **Main Application** (`app/main.py`) - Entry point
- **Tower Bolt Package** - Core functionality modules
  - `gui.py` - User interface
  - `funcs.py` - Data processing functions
  - `flange.py` - Flange calculations
  - `reporting.py` - PDF report generation
- **Report Template** - Excel template for data formatting
- **Configuration** - Customizable settings

## Dependencies

The application uses:
- PyQt5 - User interface
- pandas - Data processing
- matplotlib - Data visualization
- openpyxl - Excel file handling
- Pillow - Image processing
- numpy - Numerical computations

All dependencies are automatically installed by the installer.

## Usage

1. Launch the application
2. Load your tension data
3. Configure analysis parameters
4. Generate reports with visualizations
5. Export professional PDF reports

## Project Structure

```
VestasTensionDataTool/
├── app/
│   ├── main.py                 # Application entry point
│   ├── requirements.txt        # Python dependencies
│   ├── tension_config.json     # Configuration file
│   └── tower_bolt_package/     # Core modules
│       ├── gui.py              # User interface
│       ├── funcs.py            # Data processing
│       ├── flange.py           # Flange calculations
│       ├── reporting.py        # Report generation
│       ├── report_template.xlsx
│       └── ...
├── Install.bat                 # Installer launcher
├── installer.ps1               # Installation script
└── README.md                   # This file
```

## Updates

To update an existing installation:

1. Double-click `Install or Update from GitHub.vbs`
2. Confirm the update when prompted
3. Latest version will be downloaded automatically

## License

Vestas Internal Tool

## Support

For issues or questions, please contact the development team or create an issue on GitHub.

## Author

Benjamin Chaney (benchaney53)
Vestas Wind Systems


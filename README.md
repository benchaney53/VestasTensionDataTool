# Tower Bolt Tension Data Tool

A professional application for analyzing and reporting tower bolt tension data for Vestas wind turbines.

## Features

- 📊 Tower bolt tension data analysis
- 📈 Visual data representation with charts
- 📄 Automated PDF report generation
- 🔧 Flange calculation utilities
- 🎨 Professional Vestas-branded output
- 💾 Configuration management

## Quick Start

### Option 1: One-Click Launcher (Recommended)

1. **Double-click** `Run Tower Bolt App.vbs`
   - This will automatically set up the application and launch it
   - First run may take longer as it installs dependencies

2. **Use the application**
   - Load your tension data
   - Configure analysis parameters
   - Generate professional reports


## Requirements

### For One-Click Launcher
- Windows 7 or later
- Python 3.8 or later (automatically detected)
- Internet connection (for first-time setup only)


## What's Included

- **Main Application** (`app/main.py`) - Entry point
- **Tower Bolt Package** - Core functionality modules
  - `gui.py` - User interface
  - `funcs.py` - Data processing functions
  - `flange.py` - Flange calculations
  - `reporting.py` - PDF report generation
- **Report Template** - Excel template for data formatting
- **Configuration** - Customizable settings
- **Run Tower Bolt App.vbs** - One-click launcher (recommended)

## Dependencies

The application uses:
- PyQt5 - User interface
- pandas - Data processing
- matplotlib - Data visualization
- openpyxl - Excel file handling
- Pillow - Image processing
- numpy - Numerical computations

All dependencies are automatically installed when you first run the application.

## Usage

1. **Launch**: Double-click `Run Tower Bolt App.vbs`
2. **Load Data**: Import your tension data files
3. **Configure**: Set analysis parameters as needed
4. **Generate**: Create reports with visualizations
5. **Export**: Save professional PDF reports

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
├── Run Tower Bolt App.vbs      # One-click launcher (recommended)
└── README.md                  # This file
```

## How It Works

### One-Click Launcher (`Run Tower Bolt App.vbs`)
- Automatically finds your Python installation
- Creates an isolated virtual environment (first run only)
- Installs all required dependencies (first run only)
- Launches the application with a professional loading screen


## License

Vestas Internal Tool

## Support

For issues or questions, please contact the development team.

## Author

Benjamin Chaney (benchaney53)
Vestas Wind Systems


# Tower Bolt Tension Data Tool

A professional application for analyzing and reporting tower bolt tension data for Vestas wind turbines. This tool provides comprehensive analysis capabilities with automated report generation and professional Vestas branding.

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

### Option 2: Manual Installation with uv (Faster)

For faster installation, you can use `uv` instead of pip:

```bash
# Install uv (if not already installed)
pip install uv

# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies with uv (10-100x faster than pip)
uv pip install -r requirements.txt

# Run the application
python main.py
```


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
  - `required_rotation.txt` - Rotation requirements data
- **Report Template** (`report_template.xlsx`) - Excel template for data formatting
- **Vestas Branding Assets** - Professional logos and icons
- **Configuration** (`tension_config.json`) - Customizable settings
- **Documentation** - README files and sample reports
- **Run Tower Bolt App.vbs** - One-click launcher (recommended)

## Dependencies

The application uses:
- **PyQt5** - User interface framework
- **pandas** - Data processing and analysis
- **matplotlib** - Data visualization and charting
- **numpy** - Numerical computations
- **openpyxl** - Excel file reading and writing

All dependencies are automatically installed when you first run the application using the VBS launcher. The application now uses `uv` for faster dependency installation (10-100x faster than pip).

## Sample Output

The package includes a sample report (`Report-Baron Winds II_A01_A1-Base_Mid1-20251001_162000.pdf`) demonstrating the professional output format with:
- Vestas branding and logos
- Comprehensive data analysis
- Visual charts and graphs
- Professional formatting

## Usage

1. **Launch**: Double-click `Run Tower Bolt App.vbs`
2. **Load Data**: Import your tension data files
3. **Configure**: Set analysis parameters as needed
4. **Generate**: Create reports with visualizations
5. **Export**: Save professional PDF reports

## Project Structure

```
Tension Data Tool/
├── app/
│   ├── main.py                 # Application entry point
│   ├── requirements.txt        # Python dependencies
│   ├── tension_config.json     # Configuration file
│   └── tower_bolt_package/     # Core modules
│       ├── __init__.py         # Package initialization
│       ├── gui.py              # User interface
│       ├── funcs.py            # Data processing
│       ├── flange.py           # Flange calculations
│       ├── reporting.py        # Report generation
│       ├── report_template.xlsx # Excel template
│       ├── required_rotation.txt # Rotation data
│       ├── README.docx         # Package documentation
│       ├── README.pdf          # Package documentation (PDF)
│       ├── Vestas_Primary_Logo_RGB.png # Vestas logo
│       └── Vestas_Icon_BlueSky01_Service-tools_RGB.png # Service icon
├── Run Tower Bolt App.vbs      # One-click launcher (recommended)
└── README.md                   # This file
```

## How It Works

### One-Click Launcher (`Run Tower Bolt App.vbs`)
- Automatically finds your Python installation
- Creates an isolated virtual environment (first run only)
- Installs all required dependencies (first run only)
- Launches the application with a professional loading screen

## Troubleshooting

### First Launch Issues
- **Slow startup**: First run installs dependencies, this is normal
- **Python not found**: Ensure Python 3.8+ is installed and in your PATH
- **Permission errors**: Run as administrator if needed

### Common Issues
- **Missing data**: Ensure your input files are in the correct format
- **Report generation fails**: Check that all required fields are populated
- **GUI not responding**: Close and restart using the VBS launcher

## License

Vestas Internal Tool

## Support

For issues or questions, please contact the development team.

## Author

Benjamin Chaney (benchaney53)
Vestas Wind Systems


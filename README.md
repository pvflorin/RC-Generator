# RC Generator

![Build Status](https://github.com/YOUR_USERNAME/YOUR_REPO/workflows/Build%20RC%20Generator%20for%20Windows%20and%20macOS/badge.svg)

A powerful application for creating Route Cards and Declarations of Conformity (COC) documents from Excel data files.

## 🚀 Quick Start

### Download Ready-to-Use Versions:
- **Windows**: Download `RC_Generator.exe` from [Actions artifacts](../../actions)
- **macOS**: Download `RC_Generator.dmg` from [Actions artifacts](../../actions)

### Features:
- ✅ **Editable Output Directory**: Choose where to save your documents
- ✅ **Excel File Integration**: Load planning and technology data from Excel
- ✅ **Route Card Generation**: Create detailed route cards for manufacturing
- ✅ **COC Document Creation**: Generate compliance documents
- ✅ **Persistent Settings**: Remembers your file locations and preferences

## 📦 Automated Builds

This project automatically builds both Windows and macOS versions using GitHub Actions:

- **Windows**: `.exe` executable 
- **macOS**: `.dmg` disk image
- **Build time**: ~10-15 minutes
- **Triggered**: On every code push

## 🛠️ Development

### Requirements:
- Python 3.9+
- Dependencies listed in `requirements.txt`

### Local Development:
```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPO.git
cd YOUR_REPO
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -r requirements.txt
python route_card_coc_app.py
```

### Building Locally:
```bash
# Install PyInstaller
pip install pyinstaller

# Windows
pyinstaller RC_Generator_Windows.spec

# macOS
pyinstaller RC_Generator.spec
```

## 📋 Usage

1. **Load Excel Files**: Import your planning and technology data
2. **Select Orders**: Choose which orders to process
3. **Set Output Directory**: Choose where to save generated documents
4. **Generate Documents**: Create Route Cards or COC documents
5. **Review Output**: Find generated files in your chosen directory

## 📁 File Structure

```
RC Generator/
├── route_card_coc_app.py          # Main application
├── requirements.txt               # Python dependencies
├── Planificare Elmet.xlsx         # Sample planning data
├── Tehnologii.xlsx                # Sample technology data
├── .github/workflows/build.yml    # GitHub Actions build configuration
└── README.md                      # This file
```

## 🔄 Getting Your Builds

### From GitHub Actions:
1. Go to the **Actions** tab
2. Click the latest successful build (green ✅)
3. Download artifacts:
   - `RC_Generator_Windows` (contains .exe)
   - `RC_Generator_macOS` (contains .dmg)

### Manual Trigger:
1. Go to **Actions** → **Build RC Generator**
2. Click **"Run workflow"**
3. Wait ~10-15 minutes for completion

## 📖 Documentation

- [GitHub Actions Build Guide](GITHUB_ACTIONS_GUIDE.md)
- [Installation Instructions](README_Installation.md)
- [Windows Build Script](build_windows.bat)

## 🆘 Support

- **Build Issues**: Check the Actions tab for error logs
- **Usage Questions**: See installation documentation
- **Feature Requests**: Open an issue

## 📄 License

This project is for internal use by Elmet Technologies.

---

**Version**: 1.0.0  
**Last Updated**: October 2025  
**Platforms**: Windows, macOS

# RC Generator - Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.5] - 2025-10-18

### Fixed
- **COC Email Attachments**: Fixed issue where COC files were not attaching to email drafts
  - **Issue**: COC success message uses "salvată în:" (feminine) while code only parsed "salvat în:" (masculine)
  - **Solution**: Added handling for both grammatical forms with fallback parsing
  - Both RC and COC files now correctly attach to email drafts

### Added
- **Version Tracking**: Added VERSION file and CHANGELOG.md for better release management
- **Change History**: Complete changelog documenting all improvements and fixes

---

## [1.1.4] - 2025-10-18

### Fixed
- **Windows UI Readability**: Fixed white text on white background issue in Windows app
  - **Issue**: Text was invisible in Windows due to color theme conflicts
  - **Solution**: Added `ensure_ui_contrast()` function that:
    - Forces "Fusion" style on Windows for consistent rendering
    - Sets explicit color palette with black text on light backgrounds
    - Applies stylesheet fallback to ensure all text is readable
  - Windows app now has readable black text on light backgrounds

---

## [1.1.3] - 2025-10-15

### Added
- **macOS Installation Guide**: Created comprehensive installation documentation
  - Added `MACOS_INSTALLATION.md` with step-by-step instructions
  - Includes multiple methods to bypass macOS Gatekeeper security warnings
  - Explains why security warnings appear and that the app is safe

### Improved
- **DMG Distribution**: Enhanced macOS .dmg packaging
  - Installation instructions included directly in DMG
  - Removed quarantine attributes where possible for smoother installation
  - Better user guidance for first-time installation

---

## [1.1.2] - 2025-10-15

### Fixed
- **macOS Build Process**: Resolved GitHub Actions build failures
  - **Issue**: App name mismatch in build verification (RC_Generator.app vs RC Generator.app)
  - **Solution**: Updated verification step to check for correct app name with underscore
  - macOS .dmg file now builds successfully

---

## [1.1.1] - 2025-10-15

### Fixed
- **GitHub Actions Build**: Fixed multiple CI/CD build issues
  - **Issue 1**: Missing RC_Generator.spec file causing "Spec file not found" error
  - **Solution**: Removed spec file dependency, using direct PyInstaller command line arguments
  - **Issue 2**: PyQt6 import errors in CI environment
  - **Solution**: Fixed import statement to use `from PyQt6.QtCore import QT_VERSION_STR`
  - **Issue 3**: Interactive prompts blocking CI pipeline
  - **Solution**: Added `--noconfirm` flag to PyInstaller
  - Both Windows .exe and macOS .dmg now build automatically via GitHub Actions

---

## [1.1.0] - 2025-10-15

### Fixed
- **Application Stability**: Resolved crashes and GUI conflicts
  - **Issue**: App crashed when processing specific orders (e.g., INR000055) due to tkinter/PyQt6 conflicts
  - **Solution**: Replaced all tkinter dialogs with PyQt6 equivalents
    - File selection dialogs now use `QFileDialog.getOpenFileName()`
    - Message boxes now use `QMessageBox.question()` and `QMessageBox.critical()`
  - Removed tkinter import from application
  - App now runs stably with system Python and PyQt6

### Added
- **Editable Output Directory**: Users can now customize where files are saved
  - Added directory selection button in GUI
  - Output path persists between sessions
  - Default to Desktop/RC_Generator_Fisiere if not set

### Improved
- **Cross-Platform Compatibility**: Better handling of different operating systems
  - Fixed PyQt6 platform plugin issues on macOS
  - Improved Windows registry and macOS config file handling
  - Enhanced error messages for missing dependencies

---

## [1.0.0] - 2025-10-10

### Added
- **Initial Release**: RC Generator application with PyQt6 GUI
  - Generate Route Cards (RC) from order data
  - Generate Certificates of Conformity (COC)
  - Batch processing support
  - Email integration for sending generated documents
  - Excel file reading and writing
  - Technology database integration
  - Windows and macOS support

### Features
- Modern PyQt6 GUI with tabbed interface
- Order selection and processing
- Automatic folder creation for each order
- Excel file generation with formatted templates
- Email draft creation with attachments (macOS Mail and Windows Outlook)
- CLI mode for headless/batch processing
- Configuration persistence (Windows registry / macOS config file)

---

## Version Numbering

- **Major version (X.0.0)**: Incompatible API changes or major new features
- **Minor version (X.Y.0)**: New functionality in a backward compatible manner
- **Patch version (X.Y.Z)**: Backward compatible bug fixes

---

## Issue Tracking Template

### Issue: [Brief Description]
- **Reported**: [Date]
- **Symptom**: [What the user experienced]
- **Root Cause**: [Technical explanation of the problem]
- **Solution**: [How it was fixed]
- **Verification**: [How to verify the fix works]

---

## Future Improvements
- Code signing for macOS to eliminate security warnings (requires Apple Developer Account)
- Automated testing suite
- User preferences panel
- Multi-language support
- Database backend for order management

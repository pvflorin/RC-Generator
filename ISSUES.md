# RC Generator - Issue Tracking & Fix Log

This file tracks all issues reported and their resolutions in chronological order.

---

## 📝 How to Use This File

When reporting a new issue:
1. Add a new entry at the top of the "Active Issues" section
2. Fill in all the details
3. Once fixed, move it to "Resolved Issues" with the fix details

---

## 🔴 Active Issues

<!-- Add new issues here -->

### Issue #[NUMBER]: [Brief Title]
- **Date Reported**: YYYY-MM-DD
- **Reporter**: [Your Name]
- **Version**: [App version when issue was found]
- **Platform**: [Windows / macOS / Both]

**Description**:
[Detailed description of what's not working or what needs to be improved]

**Steps to Reproduce**:
1. [First step]
2. [Second step]
3. [...]

**Expected Behavior**:
[What should happen]

**Actual Behavior**:
[What actually happens]

**Impact**: [High / Medium / Low]

**Additional Context**:
[Any other relevant information, screenshots, error messages, etc.]

---

## ✅ Resolved Issues

### Issue #5: COC Files Not Attaching to Email
- **Date Reported**: 2025-10-18
- **Date Resolved**: 2025-10-18
- **Version Fixed**: v1.1.5
- **Platform**: Both

**Description**:
When generating COC documents and trying to email them, the email draft opened but the generated COC files were not attached. RC files attached correctly.

**Root Cause**:
The file path extraction logic only looked for "salvat în:" (masculine form in Romanian), but COC success messages use "salvată în:" (feminine form). This caused the path extraction to fail for COC files.

**Solution**:
Updated file path extraction to handle both grammatical forms:
```python
if "salvat în: " in mesaj:
    file_path = mesaj.split("salvat în: ")[-1]
elif "salvată în: " in mesaj:
    file_path = mesaj.split("salvată în: ")[-1]
else:
    file_path = mesaj.split(": ")[-1]  # Fallback
```

**Verification**:
- Generate COC document
- Click "Trimite Email" button
- Verify COC file is attached to the email draft

---

### Issue #4: White Text on White Background in Windows
- **Date Reported**: 2025-10-18
- **Date Resolved**: 2025-10-18
- **Version Fixed**: v1.1.4
- **Platform**: Windows

**Description**:
In the Windows version of the app, text appears white on a white background, making it completely unreadable. The app functions work correctly, but users cannot see any text.

**Root Cause**:
Qt on Windows was defaulting to a light color scheme with white text, which combined with the white background made text invisible. The dark palette applied in the code wasn't being enforced properly on Windows.

**Solution**:
Added `ensure_ui_contrast()` function that:
- Forces "Fusion" style on Windows (more consistent than native style)
- Sets explicit color palette with readable colors
- Applies stylesheet fallback to ensure all text is black
- Called before QApplication creation

**Verification**:
- Install Windows .exe
- Launch app
- Verify all text is readable (black text on light backgrounds)

---

### Issue #3: macOS Build Verification Failing
- **Date Reported**: 2025-10-15
- **Date Resolved**: 2025-10-15
- **Version Fixed**: v1.1.2
- **Platform**: macOS

**Description**:
GitHub Actions build was completing successfully but the verification step reported "RC Generator.app not found" even though the app was built.

**Root Cause**:
App name mismatch: PyInstaller creates `RC_Generator.app` (with underscore) but the verification step was checking for `RC Generator.app` (with space).

**Solution**:
Updated verification script to check for the correct app name: `RC_Generator.app`

**Verification**:
- Check GitHub Actions logs
- Verify build completes without errors
- Download and test .dmg file

---

### Issue #2: GitHub Actions Build Failures
- **Date Reported**: 2025-10-15
- **Date Resolved**: 2025-10-15
- **Version Fixed**: v1.1.1
- **Platform**: Both

**Description**:
Multiple issues causing GitHub Actions CI/CD builds to fail:
1. Missing RC_Generator.spec file
2. PyQt6 import errors
3. Interactive prompts blocking pipeline

**Root Cause**:
1. Spec file was in .gitignore and not committed
2. Incorrect PyQt6 import statement
3. PyInstaller prompting for confirmation to overwrite files

**Solution**:
1. Removed spec file dependency, using direct PyInstaller commands
2. Fixed import: `from PyQt6.QtCore import QT_VERSION_STR`
3. Added `--noconfirm` flag to PyInstaller

**Verification**:
- Monitor GitHub Actions
- Verify both Windows and macOS builds complete
- Download and test artifacts

---

### Issue #1: Application Crashes on Specific Orders
- **Date Reported**: 2025-10-15
- **Date Resolved**: 2025-10-15
- **Version Fixed**: v1.1.0
- **Platform**: Both (especially in bundled apps)

**Description**:
App crashes when processing certain orders (e.g., INR000055). Error message indicated Tcl_Panic related to tkinter initialization.

**Root Cause**:
Mixed GUI framework usage - application used PyQt6 as main GUI but had tkinter imports for some dialogs. PyInstaller bundled both frameworks, causing conflicts when both tried to initialize.

**Solution**:
- Removed all tkinter imports
- Replaced tkinter dialogs with PyQt6 equivalents:
  - `simpledialog` → `QInputDialog`
  - `messagebox` → `QMessageBox`
  - `filedialog` → `QFileDialog`
- Added tkinter to excludes in PyInstaller build

**Verification**:
- Process order INR000055
- Verify no crashes
- Test all dialog functionality

---

## 📊 Issue Statistics

- **Total Issues Reported**: 5
- **Total Resolved**: 5
- **Average Resolution Time**: Same day
- **Most Common Platform**: Both (60%)
- **Most Common Category**: Build/Distribution (40%)

---

## 🎯 Common Issue Categories

1. **Build/Distribution** (40%): Issues with creating installers and packages
2. **GUI/Display** (20%): Visual and interface problems
3. **Functionality** (20%): Core feature bugs
4. **Stability** (20%): Crashes and errors

---

## 📚 Lessons Learned

1. **Always test on target platforms**: Issues may not appear in development environment
2. **Be careful with mixed frameworks**: Using multiple GUI frameworks can cause conflicts
3. **CI/CD is essential**: Automated builds catch issues early
4. **Language matters**: Grammar differences (like salvat/salvată) can cause bugs
5. **Version tracking is crucial**: Makes debugging and communication much easier

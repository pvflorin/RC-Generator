# Documentation System

This folder contains comprehensive documentation for the RC Generator project.

## 📄 Documentation Files

### VERSION
- Current version number (e.g., `1.1.5`)
- Updated with each release
- Used by the app to display version in title bar

### CHANGELOG.md
- **Complete history** of all changes to the project
- Organized by version in reverse chronological order
- Includes Added, Fixed, and Improved sections
- Follows [Keep a Changelog](https://keepachangelog.com/) format
- Use this to understand what changed between versions

### ISSUES.md
- **Active issue tracking** and resolution log
- Template for reporting new issues
- Complete history of resolved issues with:
  - Detailed problem descriptions
  - Root cause analysis
  - Solution implemented
  - Verification steps
- Use this to:
  - Report new bugs or feature requests
  - Track progress on open issues
  - Learn from past fixes

### MACOS_INSTALLATION.md
- Detailed installation guide for macOS users
- Solutions for security warnings
- Troubleshooting tips
- Use this when distributing to macOS users

### README_Installation.md
- General installation instructions
- Cross-platform guidance
- Use this for Windows users or general distribution

---

## 🔄 Workflow for Updates

### When Making Changes:

1. **Fix the code**
   - Make necessary changes to `route_card_coc_app.py` or other files
   - Test thoroughly on target platforms

2. **Update VERSION**
   ```bash
   # Increment version number
   echo "1.1.6" > VERSION
   ```
   Version numbering:
   - Patch (X.X.Z): Bug fixes
   - Minor (X.Y.0): New features
   - Major (X.0.0): Breaking changes

3. **Update CHANGELOG.md**
   - Add new section at the top with current version and date
   - Document what was Added, Fixed, or Improved
   - Be specific about what changed and why

4. **Update ISSUES.md**
   - If fixing a bug: Move issue from Active to Resolved
   - Add root cause and solution details
   - If adding feature: Document the enhancement

5. **Commit with version tag**
   ```bash
   git add .
   git commit -m "v1.1.6 - [Brief description]"
   git tag -a v1.1.6 -m "Version 1.1.6 - [Description]"
   git push origin main
   git push origin v1.1.6
   ```

6. **Create GitHub Release** (optional)
   - Go to GitHub → Releases
   - Create release from tag
   - Copy relevant CHANGELOG section to release notes

---

## 📋 Issue Reporting Template

When reporting a new issue, use this format in ISSUES.md:

```markdown
### Issue #[NUMBER]: [Brief Title]
- **Date Reported**: YYYY-MM-DD
- **Reporter**: [Name]
- **Version**: [e.g., v1.1.5]
- **Platform**: [Windows / macOS / Both]

**Description**:
[What's wrong or what needs improvement]

**Steps to Reproduce**:
1. [Step 1]
2. [Step 2]

**Expected Behavior**:
[What should happen]

**Actual Behavior**:
[What actually happens]

**Impact**: [High / Medium / Low]
```

---

## 📈 Version History Quick Reference

| Version | Date | Key Changes |
|---------|------|-------------|
| 1.1.5 | 2025-10-18 | Version tracking system, changelog, issue tracking |
| 1.1.4 | 2025-10-18 | Windows UI contrast fix |
| 1.1.3 | 2025-10-15 | macOS installation guide, DMG improvements |
| 1.1.2 | 2025-10-15 | macOS build verification fix |
| 1.1.1 | 2025-10-15 | GitHub Actions build fixes |
| 1.1.0 | 2025-10-15 | Stability fixes, editable output directory |
| 1.0.0 | 2025-10-10 | Initial release |

---

## 🎯 Best Practices

1. **Always update VERSION file** when making changes
2. **Keep CHANGELOG current** - update it with every commit
3. **Track issues properly** - move them from Active to Resolved
4. **Be detailed** - future you will thank present you
5. **Use version tags** - makes it easy to find specific releases
6. **Test before tagging** - tags should represent stable versions

---

## 🔗 Related Files

- `.github/workflows/build.yml` - Automated build configuration
- `route_card_coc_app.py` - Main application code
- `requirements.txt` - Python dependencies

---

## 💡 Tips

- Search ISSUES.md for similar problems before reporting new ones
- Check CHANGELOG.md to see if your issue was already fixed
- Use GitHub's issue tracker for public bugs
- Use ISSUES.md for internal tracking during development

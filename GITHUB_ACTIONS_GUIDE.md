# RC Generator - GitHub Actions Build Guide

## 🚀 Automated Building with GitHub Actions

This project is set up to automatically build both **Windows (.exe)** and **macOS (.dmg)** versions using GitHub Actions. No need for manual building!

## 📋 Setup Instructions

### 1. Create GitHub Repository
1. Go to [GitHub.com](https://github.com) and create a new repository
2. Name it something like `RC-Generator` or `route-card-generator`
3. Make it **Public** (required for free GitHub Actions)

### 2. Upload Your Code
```bash
# From your project directory
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
git branch -M main
git push -u origin main
```

### 3. Automatic Building
Once you push to GitHub, the builds will automatically start:

- **Windows build**: Creates `RC_Generator.exe`
- **macOS build**: Creates `RC_Generator.dmg`
- **Build time**: ~10-15 minutes for both platforms

## 📦 Downloading Your Builds

### Method 1: GitHub Actions Artifacts
1. Go to your GitHub repository
2. Click **"Actions"** tab
3. Click on the latest successful build (green checkmark ✅)
4. Scroll down to **"Artifacts"** section
5. Download:
   - `RC_Generator_Windows` - Contains the .exe file
   - `RC_Generator_macOS` - Contains the .dmg file

### Method 2: GitHub Releases (For Tagged Versions)
1. Create a new release/tag in GitHub
2. The builds will automatically attach to the release
3. Users can download directly from the Releases page

## 🔄 Triggering New Builds

Builds automatically trigger when you:
- **Push code** to the main branch
- **Create pull requests**
- **Manually trigger** from Actions tab

### Manual Trigger:
1. Go to **Actions** tab in GitHub
2. Click **"Build RC Generator for Windows and macOS"**
3. Click **"Run workflow"**
4. Select branch and click **"Run workflow"**

## 📁 What Gets Built

### Windows Package (`RC_Generator_Windows`):
```
RC_Generator_Windows_Installer/
├── RC_Generator.exe          # Main executable
├── README.txt               # Installation instructions
├── Planificare Elmet.xlsx   # Sample data file
└── Tehnologii.xlsx          # Sample data file
```

### macOS Package (`RC_Generator_macOS`):
```
RC_Generator_macOS_Installer/
├── RC_Generator.dmg         # Disk image installer
└── README_Installation.md   # Installation instructions
```

## 🛠️ Customizing Builds

### Adding New Dependencies:
Edit `requirements.txt` and commit changes:
```txt
pandas==2.0.3
numpy==1.26.4
your-new-package==1.0.0
```

### Changing Build Settings:
Edit `.github/workflows/build.yml`:
- Add new build steps
- Change Python version
- Modify PyInstaller options
- Add code signing (advanced)

## 🔧 Troubleshooting

### Build Fails?
1. Check the **Actions** tab for error logs
2. Common issues:
   - Missing dependencies in `requirements.txt`
   - Python version compatibility
   - File path issues

### Downloads Not Working?
- Ensure repository is **Public** (for free Actions)
- Check that build completed successfully (green checkmark)
- Artifacts expire after 90 days

### Need Different Python Version?
Edit `.github/workflows/build.yml`:
```yaml
- name: Set up Python
  uses: actions/setup-python@v4
  with:
    python-version: '3.11'  # Change here
```

## 📊 Build Status

You can add a build status badge to your README:
```markdown
![Build Status](https://github.com/YOUR_USERNAME/YOUR_REPO/workflows/Build%20RC%20Generator%20for%20Windows%20and%20macOS/badge.svg)
```

## 🎯 Next Steps

1. **Push your code to GitHub**
2. **Wait for builds to complete** (~10-15 minutes)
3. **Download your executables** from Actions artifacts
4. **Distribute** to users as needed

## 💡 Pro Tips

- **Tag releases** for permanent download links
- **Enable notifications** to know when builds complete
- **Star your repo** to make it easier to find
- **Add collaborators** if working in a team

---

**Need help?** Check the GitHub Actions documentation or the build logs in your repository's Actions tab.

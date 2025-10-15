# RC Generator - macOS Installation Guide

## 🍎 Installing on macOS

### Step 1: Download
Download the `RC_Generator.dmg` file from the GitHub Actions artifacts or releases.

### Step 2: Mount the DMG
Double-click the `RC_Generator.dmg` file to mount it.

### Step 3: Install the App
Drag `RC_Generator.app` to the Applications folder.

### Step 4: Handle Security Warning
When you first try to open the app, macOS will show a security warning because the app isn't signed with an Apple Developer certificate.

**Method 1 (Recommended):**
1. Right-click on `RC_Generator.app` in Applications
2. Select "Open" from the context menu
3. Click "Open" when prompted about the unverified developer

**Method 2:**
1. Go to **System Preferences** → **Security & Privacy** → **General**
2. Look for "RC_Generator was blocked" message
3. Click **"Open Anyway"**

**Method 3 (Advanced Users):**
```bash
sudo xattr -rd com.apple.quarantine /Applications/RC_Generator.app
```

### Step 5: Normal Usage
After the first successful open, macOS will remember your choice and the app will open normally.

## 🔒 Why This Warning Appears

This security warning appears because:
- The app is not signed with an Apple Developer certificate ($99/year)
- macOS Gatekeeper blocks unsigned apps by default
- **The app is completely safe** - it's just a precautionary measure

## 🛠 Troubleshooting

If you continue to have issues:
1. Make sure you're running macOS 10.14 or later
2. Ensure you have admin privileges on your Mac
3. Try restarting your Mac after installation
4. Check that the app is in the Applications folder, not running from the DMG

## 📧 Support

If you encounter any issues, please report them in the GitHub repository issues section.

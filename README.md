# Hebrew Fonts Ribbon for Microsoft Word

🎨 A custom ribbon add-in for Microsoft Word that organizes Hebrew fonts (BA, BU, FB, LIA, EFT) with favorites, recent fonts, and family browsing.
 [![Download](https://img.shields.io/badge/Download-DOTM-blue)](https://github.com/rabbieliazer/hebrew-fonts-ribbon/raw/main/hebrew%20fonts.dotm)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Word](https://img.shields.io/badge/Word-2010%2B-orange)](https://www.microsoft.com/word)

## ✨ Features

- ⭐ **Favorites System** - Star your most-used fonts for quick access
- 📝 **Recent Fonts** - Automatic tracking of your last 10 fonts
- 📚 **Company Groups** - Separate dropdowns for BA, BU, FB, LIA, and EFT
- 👨‍👩‍👧‍👦 **Family Browser** - Browse by font family and select variants
-  💾 **Persistent Storage** - Your favorites and recents are saved between sessions
- 🎯 **Current Font Display** - See which font is currently selected
- 🌟 **One-Click Favorites** - Star button next to each company dropdown
- 🗑️ **Easy Uninstall** - Built-in uninstall button with step-by-step guidance

## 📋 Requirements

- **Microsoft Word 2010 or later** (Windows)
- **Hebrew fonts installed** (BA, BU, FB, LIA, or EFT families)
- **Windows 7 or later**

## 🚀 Installation

### ⚡ Method 1: One-Click Installer EXE (Easiest - Recommended)

1. **[Download HebrewFontsInstaller.exe](https://github.com/rabbieliazer/hebrew-fonts-ribbon/raw/main/HebrewFontsInstaller.exe)**
2. **Double-click** `HebrewFontsInstaller.exe` to run
3. **Follow the on-screen prompts**
4. **Restart Microsoft Word**

The installer automatically:
- ✅ Checks if Word is running
- ✅ Downloads the latest version
- ✅ Installs to the correct folder
- ✅ Cleans up temporary files

> **Note**: Your browser or Windows may show a security warning because this is a new file. Click "Keep" or "Run anyway" - the file is safe and open source.

***

### ⚡ Method 2: Batch File Installer (Alternative)

1. **[Download install.bat](https://github.com/rabbieliazer/hebrew-fonts-ribbon/raw/main/install.bat)**
2. **Double-click** `install.bat` to run
3. **Follow the prompts**
4. **Restart Microsoft Word**

***

### 📥 Method 3: Manual Installation

#### Step 1: Download
**[Download hebrew fonts.dotm](https://github.com/rabbieliazer/hebrew-fonts-ribbon/raw/main/hebrew%20fonts.dotm)**

#### Step 2: Enable Macros
1. Open **Microsoft Word**
2. Go to **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. Click **Macro Settings**
4. Select **"Enable all macros"**
5. Click **OK** twice

#### Step 3: Install the Add-in
1. Press `Win + R` on your keyboard
2. Type: `%APPDATA%\Microsoft\Word\STARTUP`
3. Press **Enter** (opens the Word STARTUP folder)
4. **Copy** `hebrew fonts.dotm` into this folder
5. **Restart Microsoft Word**

#### Step 4: Verify Installation
Look for the **"Hebrew Fonts"** tab in your Word ribbon ✅

***

### 🔐 Security Notice (First Time Only)

When you first open Word, you might see a security warning:

1. Click **Enable Content** or **Enable Macros**
2. If you see this warning every time:
   - File → Options → Trust Center → **Trusted Locations**
   - Click **Add new location**
   - Browse to: `%APPDATA%\Microsoft\Word\STARTUP`
   - Check **"Subfolders of this location are also trusted"**
   - Click **OK**
***

### 🔐 Security Notice (First Time Only)

When you first open Word, you might see a security warning:

1. Click **Enable Content** or **Enable Macros**
2. If you see this warning every time:
   - File → Options → Trust Center → **Trusted Locations**
   - Click **Add new location**
   - Browse to: `%APPDATA%\Microsoft\Word\STARTUP`
   - Check **"Subfolders of this location are also trusted"**
   - Click **OK**

***

## 📖 How to Use

### Quick Access
- **Favorites Dropdown**: Access your starred fonts instantly
- **Recent Dropdown**: See your last 10 used fonts
- **Current Font**: View the currently selected font in your document

### Select Fonts by Company
1. Click a company dropdown (**BA**, **BU**, **FB**, **LIA**, or **EFT**)
2. Choose your font
3. Font is **automatically applied** to selected text
4. Click the **⭐ star button** to add it to favorites

### Browse by Family
1. Select **Company** from the first dropdown
2. Select **Family** (e.g., "David", "Frank", "Rashi")
3. Select **Variant** (Bold, Regular, Light, etc.)
4. Click **⭐** to add to favorites

### Manage Your Favorites
- **Add Current Font**: Click "Favorite Current" to save the font you're using now
- **Manage Favorites**: View all favorites and remove any (max 20)
- **Clear Recent**: Reset your recent fonts list

***

## 🗑️ Uninstall

### Method 1: Built-in Uninstaller (Easiest)
1. Open **Microsoft Word**
2. Go to the **Hebrew Fonts** tab
3. Click the **Uninstall** button (Settings group)
4. Follow the step-by-step instructions
5. Delete the file when File Explorer opens
6. **Restart Word**

### Method 3: Manual Uninstall
1. Press `Win + R`
2. Type: `%APPDATA%\Microsoft\Word\STARTUP`
3. Press **Enter**
4. Delete `hebrew fonts.dotm`
5. **Restart Word**

***

## 🔄 Updating to a Newer Version

**Option 1**: Run `install.bat` again (automatically overwrites old version)

**Option 2**: 
1. Uninstall the old version
2. Download and install the new version

***

## 🛠️ Troubleshooting

### ❌ Ribbon Doesn't Appear
- ✅ Ensure file is in `%APPDATA%\Microsoft\Word\STARTUP`
- ✅ Restart Word completely (close all windows)
- ✅ Check that macros are enabled
- ✅ Try reinstalling with `install.bat`

### ⚠️ Security Warning Every Time
1. File → Options → Trust Center → **Trusted Locations**
2. Click **Add new location**
3. Browse to STARTUP folder: `%APPDATA%\Microsoft\Word\STARTUP`
4. Check **"Subfolders of this location are also trusted"**
5. Click **OK**

### 📁 Can't Find STARTUP Folder
1. Press `Win + R`
2. Copy and paste: `%APPDATA%\Microsoft\Word\STARTUP`
3. Press **Enter**
4. If folder doesn't exist, create it manually

### 🔒 Browser Blocks Download
- Your browser may show a warning
- Click **"Keep"** or **"Download anyway"**
- The file is safe (open source and verified)

### 🚫 Fonts Not Showing
- ✅ Verify Hebrew fonts are installed in Windows
- ✅ Font names must start with: BA, BU, FB, LIA, or EFT
- ✅ Restart Word to refresh font cache

### 💥 Errors on Startup
1. Delete the file from STARTUP folder
2. Download a fresh copy
3. Copy back to STARTUP folder
4. Restart Word

***

## 🎯 Supported Font Families

The ribbon automatically detects these font families:

| Company | Pattern | Examples |
|---------|---------|----------|
| **BA** | Starts with "BA " or "BAWS " | BA David, BAWS Miriam |
| **BU** | Starts with "BU " | BU Frank, BU Rashi |
| **FB** | Starts with "FB " | FB Narkis, FB Hadassa |
| **LIA** | Starts with "LIA " | LIA David, LIA Rashi |
| **EFT** | Starts with "EFT " | EFT Frank, EFT Miriam |

***

## 🔧 Advanced Configuration

### Change Default Company

The default company in the Family Browser is **BA**. To change:
1. Open VBA Editor in Word (`Alt + F11`)
2. Find: `selectedCompany = "BA"`
3. Change to: `selectedCompany = "BU"` (or your preferred company)
4. Save and restart Word

### Adjust Maximum Favorites

Default limit is **20 favorites**. To change:
1. Open VBA Editor (`Alt + F11`)
2. Find: `If favoriteFonts.Count >= 20 Then`
3. Change `20` to your desired number
4. Save and restart Word

### Installation Locations

**Default** (recommended):

%APPDATA%\Microsoft\Word\STARTUP

**Alternative** (requires admin rights):

%ProgramFiles%\Microsoft Office\root\Office16\STARTUP

***

## 📞 Support & Feedback

- 🐛 **Report Bugs**: [Open an issue](https://github.com/rabbieliazer/hebrew-fonts-ribbon/issues)
- 💬 **Questions**: [Start a discussion](https://github.com/rabbieliazer/hebrew-fonts-ribbon/discussions)
- ⭐ **Like it?**: [Star this repository!](https://github.com/rabbieliazer/hebrew-fonts-ribbon)

***

## 📝 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

**In short**: Free to use, modify, and distribute. No warranty provided.

***

## 🙏 Acknowledgments

- Created with ❤️ for the Hebrew typography community by **Tosh Developers**
- Thanks to Hebrew font foundries: BA, BU, FB, LIA, EFT
- Built with Microsoft Office RibbonX framework
- Special thanks to all contributors and users

***

## 📊 Project Information

- **Version**: 1.1
- **Last Updated**: March 2026
- **Author**: Tosh Developers
- **Repository**: [github.com/rabbieliazer/hebrew-fonts-ribbon](https://github.com/rabbieliazer/hebrew-fonts-ribbon)
- **Compatible With**: Word 2010, 2013, 2016, 2019, 2021, 365 (Windows)

***

## 🌟 Show Your Support

If you find this useful, please **⭐ star this repository** to help others discover it!

[![GitHub stars](https://img.shields.io/github/stars/rabbieliazer/hebrew-fonts-ribbon?style=social)](https://github.com/rabbieliazer/hebrew-fonts-ribbon/stargazers)

***

## 🔄 Changelog

### Version 1.1 (Current)
- ✨ Added built-in uninstall button with step-by-step guidance
- 🐛 Fixed font detection for short font names
- ⚡ Improved font cache initialization performance
- 📚 Extended font style recognition (22 styles)
- 🧹 Auto-cleanup of favorites and recent fonts on uninstall

### Version 1.0
- 🎉 Initial release
- ⭐ Favorites system
- 📝 Recent fonts tracking
- 📚 Company-based font organization
- 👨‍👩‍👧‍👦 Family browser

***

**Made with ❤️ by Tosh Developers**

*For the Hebrew typography community*

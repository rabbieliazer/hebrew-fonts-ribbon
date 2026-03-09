# Hebrew Fonts Ribbon for Microsoft Word

🎨 A custom ribbon add-in for Microsoft Word that organizes Hebrew fonts (BA, BU, FB, LIA, EFT) with favorites, recent fonts, and family browsing.

[![Download](https://img.shields.io/badge/Download-DOTM-blue)](https://github.com/rabbieliazer/hebrew-fonts-ribbon/raw/main/hebrew%20fonts.dotm)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE) [![Word](https://img.shields.io/badge/Word-2010%2B-orange)](https://www.microsoft.com/word)

## ✨ Features

- ⭐ **Favorites System** - Star your most-used fonts for quick access
- 📝 **Recent Fonts** - Automatic tracking of your last 10 fonts
- 📚 **Company Groups** - Separate dropdowns for BA, BU, FB, LIA, and EFT
- 👨‍👩‍👧‍👦 **Family Browser** - Browse by font family and select variants
- 💾 **Persistent Storage** - Your favorites and recents are saved between sessions
-  🎯 **Current Font Display** - See which font is currently selected
- 🌟 **One-Click Favorites** - Star button next to each company dropdown

## 📋 Requirements

- **Microsoft Word 2010 or later** (Windows)
- **Hebrew fonts installed** (BA, BU, FB, LIA, or EFT families)
- **Windows 7 or later**

## 🚀 Installation

### ⚡ Method 1: One-Click Installer (Easiest - Recommended)

1. **Download the installer**: [`install.bat`](https://github.com/rabbieliazer/hebrew-fonts-ribbon/raw/main/install.bat) 2. **Right-click** `install.bat` → **Run as administrator** (or just double-click)
3. **Follow the prompts**
4. **Restart Microsoft Word**

That's it! The installer will:
- ✅ Check if Word is running
- ✅ Download the latest version
- ✅ Install to the correct folder
- ✅ Clean up temporary files

---

### 📥 Method 2: Manual Installation

#### Step 1: Download
Click here to download: **[hebrew fonts.dotm](https://github.com/rabbieliazer/hebrew-fonts-ribbon/raw/main/hebrew%20fonts.dotm)**

#### Step 2: Enable Macros
1. Open **Microsoft Word**
2. Go to **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. Click **Macro Settings**
4. Select **"Enable all macros"** 
5. Click **OK** twice

#### Step 3: Install the Add-in
1. Press `Win + R` on your keyboard
2. Type: `%APPDATA%\Microsoft\Word\STARTUP`
3. Press **Enter** (this opens the Word STARTUP folder)
4. **Copy** `hebrew fonts.dotm` into this folder
5. **Restart Microsoft Word**

#### Step 4: Verify
Look for the **"Hebrew Fonts"** tab in your Word ribbon - you're done! 🎉

---

### 🔐 Security Notice (First Time Only)

When you first open Word, you might see a security warning:

1. Click **Enable Content** or **Enable Macros**
2. If you see this every time:
   - File → Options → Trust Center → **Trusted Locations**
   - Click **Add new location**
   - Browse to: `%APPDATA%\Microsoft\Word\STARTUP`
   - Check **"Subfolders of this location are also trusted"**
   - Click **OK**

---

## 📖 How to Use

### Quick Access Fonts
- **Favorites**: Click the favorites dropdown to access your starred fonts
- **Recent**: See your last 10 used fonts in the recent dropdown
- **Current Font**: See which font is currently selected in your document

### Select Fonts by Company
1. Click a company dropdown (**BA**, **BU**, **FB**, **LIA**, or **EFT**)
2. Choose your font
3. The font is **automatically applied** to selected text
4. Click the **⭐ star button** next to the dropdown to add it to favorites

### Browse by Family
1. Select **Company** from the first dropdown
2. Select **Family** (e.g., "David", "Frank", "Rashi")
3. Select **Variant** (Bold, Regular, Light, etc.)
4. Click **⭐** to favorite it

### Manage Favorites
- **Add Current Font**: Click "Favorite Current" to add the font you're using now
- **Manage Favorites**: Click to view all favorites and remove any (max 20)
- **Clear Recent**: Reset your recent fonts list

---

## 🛠️ Troubleshooting

### ❌ Ribbon doesn't appear
- ✅ Make sure the file is in `%APPDATA%\Microsoft\Word\STARTUP`
- ✅ Restart Word **completely** (close all windows)
- ✅ Check that macros are enabled
- ✅ Try reinstalling using `install.bat`

### ⚠️ Security Warning Every Time
1. File → Options → Trust Center → **Trusted Locations**
2. Click **Add new location**
3. Browse to your STARTUP folder
4. Check **"Subfolders of this location are also trusted"**
5. Click **OK**

### 📁 Can't Find STARTUP Folder
1. Press `Win + R`
2. Type exactly: `%APPDATA%\Microsoft\Word\STARTUP`
3. Press Enter
4. If folder doesn't exist, create it manually

### 🔒 File Won't Download
- ✅ Check your internet connection
- ✅ Try the direct link: [Click here](https://github.com/rabbieliazer/hebrew-fonts-ribbon/raw/main/hebrew%20fonts.dotm)
- ✅ Your browser might block it - click "Keep" or "Download anyway"

### 🚫 Fonts Not Showing
- ✅ Verify your Hebrew fonts are installed in Windows
- ✅ Font names must start with: BA, BU, FB, LIA, or EFT
- ✅ Restart Word to refresh the font cache

### 💥 Errors on Load
1. Delete the file from STARTUP folder
2. Download a fresh copy
3. Right-click the file → **Properties** → Check **"Unblock"** → **OK**
4. Copy back to STARTUP folder

---

## 🗑️ Uninstall

### Using Uninstaller (Easy)
1. Download [`uninstall.bat`](https://github.com/rabbieliazer/hebrew-fonts-ribbon/raw/main/uninstall.bat)
2. Double-click to run
3. Confirm removal
4. Restart Word

### Manual Uninstall
1. Press `Win + R`
2. Type: `%APPDATA%\Microsoft\Word\STARTUP`
3. Delete `hebrew fonts.dotm`
4. Restart Word

---

## 🔄 Updating

To update to a newer version:

**Option 1**: Run `install.bat` again (it will overwrite)

**Option 2**: 
1. Uninstall the old version
2. Download and install the new version

---

## 📞 Support & Feedback

- 🐛 **Report Bugs**: [Open an issue](https://github.com/rabbieliazer/hebrew-fonts-ribbon/issues)
- 💬 **Questions**: [Start a discussion](https://github.com/rabbieliazer/hebrew-fonts-ribbon/discussions)
-  ⭐ **Like it?**: Star this repository!

---

## 🎯 Supported Font Families

The ribbon automatically detects these font families:

| Company | Pattern | Examples |
|---------|---------|----------|
| **BA** | Starts with "BA " or "BAWS " | BA David, BAWS Miriam |
| **BU** | Starts with "BU " | BU Frank, BU Rashi |
| **FB** | Starts with "FB " | FB Narkis, FB Hadassa |
| **LIA** | Starts with "LIA " | LIA David, LIA Rashi |
| **EFT** | Starts with "EFT " | EFT Frank, EFT Miriam |

---

##  🔧 Advanced

### Change Installation Location

By default, the add-in installs to:

%APPDATA%\Microsoft\Word\STARTUP

You can also install to:

%ProgramFiles%\Microsoft Office\root\Office16\STARTUP

(Requires administrator rights)

### Customize Default Company

To change the default company in the Family Browser:
1. Open VBA Editor in Word (`Alt + F11`)
2. Find: `selectedCompany = "BA"`
3. Change to: `selectedCompany = "BU"` (or any other company)

### Maximum Favorites

Default limit is 20 favorites. To change:
1. Open VBA Editor
2. Find: `If favoriteFonts.Count >= 20 Then`
3. Change `20` to your desired number

---

## 📝 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

**In short**: Free to use, modify, and distribute. No warranty provided.

---

## 🙏 Acknowledgments

- Created with ❤️ for the Hebrew typography community
- Thanks to all Hebrew font foundries (BA, BU, FB, LIA, EFT)
- Built with Microsoft Office RibbonX framework

---

## 📊 Project Info

- **Version**: 1.0
- **Last Updated**: March 2026
- **Author**: Rabbi Eliazer
- **Compatible With**: Word 2010, 2013, 2016, 2019, 2021, 365

---

## 🌟 Star This Repository

If you find this useful, please **star** this repository to help others discover it!

[![GitHub stars](https://img.shields.io/github/stars/rabbieliazer/hebrew-fonts-ribbon?style=social)](https://github.com/rabbieliazer/hebrew-fonts-ribbon)

---

## 📸 Screenshots

*(Add screenshots here when available)*

---

**Made with ❤️ by Rabbi Eliazer**
 *For the Hebrew typography community*

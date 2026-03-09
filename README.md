# Hebrew Fonts Ribbon for Microsoft Word

🎨 A custom ribbon add-in for Microsoft Word that organizes Hebrew fonts (BA, BU, FB, LIA, EFT) with favorites, recent fonts, and family browsing.

## ✨ Features

- ⭐ **Favorites System** - Star your most-used fonts for quick access
- 📝 **Recent Fonts** - Automatic tracking of your last 10 fonts
- 📚 **Company Groups** - Separate dropdowns for BA, BU, FB, LIA, and EFT
- 👨‍👩‍👧‍👦 **Family Browser** - Browse by font family and select variants
- 💾 **Persistent Storage** - Your favorites and recents are saved between sessions
- 🎯 **Current Font Display** - See which font is currently selected

## 📋 Requirements

- Microsoft Word 2010 or later (Windows)
- Hebrew fonts installed (BA, BU, FB, LIA, or EFT families)

## 🚀 Quick Installation

### Step 1: Download
Click the **green "Code" button** above → **Download ZIP** → Extract the file

OR
 Download `HebrewFontsRibbon.dotm` directly from [Releases](../../releases)

### Step 2: Enable Macros
1. Open Microsoft Word
2. Go to **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. Click **Macro Settings**
4. Select **"Enable all macros"** 
5. Click **OK**

### Step 3: Install the Add-in
1. Press `Win + R` on your keyboard
2. Type: `%APPDATA%\Microsoft\Word\STARTUP`
3. Press **Enter** (this opens the Word STARTUP folder)
4. Copy `HebrewFontsRibbon.dotm` into this folder
5. **Restart Microsoft Word**

### Step 4: Verify Look for the **"Hebrew Fonts"** tab in your Word ribbon - you're done! 🎉

## 📖 How to Use

### Quick Access Fonts
- **Favorites**: Click the favorites dropdown to access your starred fonts
- **Recent**: See your last 10 used fonts in the recent dropdown

### Select Fonts by Company
1. Click a company dropdown (BA, BU, FB, LIA, or EFT)
2. Choose your font
3. The font is automatically applied to selected text
4. Click the ⭐ star button to add it to favorites

### Browse by Family
1. Select **Company** from the first dropdown
2. Select **Family** (e.g., "David", "Frank")
3. Select **Variant** (Bold, Regular, Light, etc.)
4. Click ⭐ to favorite it

### Manage Favorites
- **Star Current Font**: Click "Favorite Current" to add the current font
- **Manage Favorites**: Click to view and remove favorites (max 20)
- **Clear Recent**: Reset your recent fonts list

## 🛠️ Troubleshooting

### Ribbon doesn't appear
- Make sure the file is in `%APPDATA%\Microsoft\Word\STARTUP`
- Restart Word completely (close all windows)
- Check that macros are enabled

### Security Warning
1. Right-click `HebrewFontsRibbon.dotm` → **Properties**
2. Check **"Unblock"** at the bottom
3. Click **OK**

### Fonts not showing
- Verify your Hebrew fonts are installed in Windows
- Restart Word to refresh the font cache

### Warning on Every Start
1. File → Options → Trust Center → **Trusted Locations**
2. Click **Add new location**
3. Browse to your STARTUP folder
4. Check **"Subfolders of this location are also trusted"**
5. Click **OK**

## 🗑️ Uninstall

1. Press `Win + R`
2. Type: `%APPDATA%\Microsoft\Word\STARTUP`
3. Delete `HebrewFontsRibbon.dotm`
4. Restart Word

## 📞 Support

- **Issues**: [Report a bug](../../issues)
- **Questions**: [Start a discussion](../../discussions)

## 📝 License

MIT License - Free to use, modify, and distribute. See [LICENSE](LICENSE) for details.

## 🙏 Acknowledgments

Created for the Hebrew typography community with  ❤️

---

**Version**: 1.0  
**Compatible with**: Microsoft Word 2010, 2013, 2016, 2019, 2021, 365  
**Author**: Rabbi Lazer

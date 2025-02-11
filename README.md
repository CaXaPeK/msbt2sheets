# Msbt2Sheets
A tool that eases the proccess of viewing the text of and (if you want) translating Nintendo games!

# Features
- Convert multiple MSBTs into a single spreadsheet
  - View all messages of a game in all languages simultaneously
  - Edit messages and convert them back to MSBTs (since it's Google Sheets, you can also invite your friends and edit together!)
  - View game script stats: total messages, total characters
  - Edit message attributes and styles
  - Add "translation columns" to all sheets and track the translation's completion
  - Add an unfinished translation to the sheets
  - Presets for fast conversion
 
- MSBP parsing
  - A list of all control tags used by the game (with their names, parameters, types, etc.)
  - A list of all styles a message can use (region width, line count, font id, base color)
  - A list of all attributes a message can set
  - A list of all base colors for coloring text
  - **Display all control tags and message attributes/styles in a human-readable form!**
 
- Formatting options
  - Shorten control tags (eg. <Text.wait msec=250> to <wait 250>)
  - Shorten the \<PageBreak\> tag to \<p\>
  - Add a line break after each \<PageBreak\> for readability
  - Drop \<Ruby\> tags in Japanese messages
  - *(Can't get it working)* Highlight control tags in **bold**

# Usage
Run Msbt2Sheets.exe and follow the on-screen instructions.

On first launch, it may crash and ask you to log into Google. You have to log in and grant access to your spreadsheets.

# Build
You'll have to provide your own Google Sheets API's clientId and clientSecret in "/data/credentials.txt". After that, build like any other C# app.

# Setting up presets
If you need to convert the same project over and over many times, you can set up a preset so you don't have to configure all the options every time. Presets are stored in "/data/presets/". Some examples are included within the release.

## Options
Keep in mind that lists are formatted like this, with a "|" separator: EUen|EUfr|EUes|EUde|JPja|KRko
### Global options
- **mode**: "1" if converting MSBTs to sheets, "2" if converting sheets to MSBTs.
- **addLinebreaksAfterPagebreaks**: true/false. Adds a line break after every \<PageBreak\> tag.

### Spreadsheet creation options
- **languagesPath**: Path to your game's language folder (the folder with EUen, EUfr, JPja and such).
- **internalLangs**: 
- **uiLangs**
- **newLangs**
- **newLangsSheetNames**
- **newLangsPaths**
- **spreadsheetName**
- **gcnTtydSpreadsheetId**
  
- **shortenTags**
- **shortenPagebreak**
- **freezeColumnCount**
- **columnSize**
- **highlightTags**
- **skipRuby**

### MSBT creation options
- **spreadsheetId**
- **outputPath**
- **uiLangs**
- **outputLangs**

- **noStatsSheet**
- **noSettingsSheet**
- **noInternalDataSheet**

- **addLinebreaksAfterPagebreaks**
- **colorIdentification**

- **skipLangIfNotTranslated**
- **extendedHeader**
- **sheetNames**
- **customFileNames**

- **globalVersion**
- **globalEndianness**
- **globalEncoding**
- **slotCounts**

- **noTranslationSymbol**
- **mainLangColumnId**

# Known bugs
- Can't work with very large games (eg. Tears of the Kingdom). *TODO: Send data in small chunks*

These early versions are extremly buggy, so if you encounter errors, please write about them in the Issues!

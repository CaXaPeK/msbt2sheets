# Msbt2Sheets

A tool that eases the proccess of viewing the text of and (if you want) translating Nintendo games!

# Features

- Convert multiple MSBTs into a single spreadsheet and back
  - View all messages of a game in all languages simultaneously
  - View game script stats: total messages, total characters
  - Edit message attributes and styles
  - Add "translation columns" to all sheets and track the translation's completion
  - Add an unfinished translation to the sheets
 
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

using System.Text.RegularExpressions;
using Google.Apis.Sheets.v4.Data;
using Msbt2Sheets.Lib.Formats;
using Msbt2Sheets.Lib.Formats.FileComponents;
using Msbt2Sheets.Lib.Utils;

namespace Msbt2Sheets.Sheets;

public class MsbtToSheets
{
    public static void Create(GoogleSheetsManager sheetsManager, Dictionary<string, string> fileOptions)
    {
        Console.Clear();
        Console.WriteLine("Enter the path to the language folder (folder with EU_English, EU_French, etc.):");
        string languagesPath;
        if (fileOptions.ContainsKey("languagesPath"))
        {
            languagesPath = fileOptions["languagesPath"].Trim('"');
        }
        else
        {
            languagesPath = Console.ReadLine().Trim('"');
        }
        ConsoleUtils.CheckDirectory(languagesPath);

        List<string> internalLangNames = ReorderLanguages(languagesPath, fileOptions);
        List<string> sheetLangNames = internalLangNames;//RenameLanguages(internalLangNames);
    
        MSBP msbp = ParseMSBP(fileOptions);
        ParsingOptions options = SetParsingOptions(fileOptions, msbp);

        options.UnnecessaryPathPrefix = FindUnnecessaryPathPrefix($"{languagesPath}/{internalLangNames[0]}/", "");
        List<List<MSBT>> languages = ParseAllMSBTs(languagesPath, internalLangNames, options, msbp);

        Spreadsheet spreadsheet = LanguagesToSpreadsheet(languages, sheetLangNames, options, fileOptions, msbp);

        if (fileOptions.ContainsKey("gcnTtydSpreadsheetId"))
        {
            HighlightTtydChanges(ref spreadsheet, sheetsManager, fileOptions["gcnTtydSpreadsheetId"]);
        }
    
        Console.WriteLine("Uploading your spreadsheet...");
        var newSheet = sheetsManager.CreateSpreadsheet(spreadsheet);
    
        Console.WriteLine($"Congratulations! Your spreadsheet is ready:\n{newSheet.SpreadsheetUrl}\n\nHave fun!");
    }
    
    static List<string> ReorderLanguages(string langPath, Dictionary<string, string> fileOptions)
    {
        List<string> langPaths = Directory.GetDirectories(langPath).ToList();
        List<string> internalLangNames = new();
        foreach (var path in langPaths)
        {
            internalLangNames.Add(Path.GetFileName(path));
        }

        if (fileOptions.ContainsKey("langs"))
        {
            internalLangNames = fileOptions["langs"].Split('|').ToList();
            return internalLangNames;
        }
        
        while (true)
        {
            Console.Clear();
            
            Console.WriteLine("This is the current order in which the languages will be exported to the spreadsheet:");
            for (int i = 0; i < internalLangNames.Count; i++)
            {
                string langName = internalLangNames[i];
                Console.Write($"{i + 1}: {langName}");
                if (i == 0) Console.Write(" (main)");
                Console.WriteLine();
            }
            Console.WriteLine("\nAre you satisfied with the order or do you want to change it?\n\n" +
                              "1 - Change order (You can also remove unwanted languages.)\n" +
                              "2 - Confirm that order\n" +
                              "3 - Choose a new main language");
            var answer = Console.ReadLine();
            if (answer == "2")
            {
                break;
            }
            if (answer == "1")
            {
                Console.WriteLine("\nEnter a new order (eg. 3 5 7 1 2). The 1st language will be considered main. You can also leave out unwanted languages.");
                var newIds = Console.ReadLine().Trim(' ').Split(' ');

                List<string> newInternalLangNames = new();
                foreach (var id in newIds)
                {
                    int intId = Convert.ToInt32(id) - 1;
                    if (intId >= internalLangNames.Count)
                    {
                        Console.Clear();
                        Console.WriteLine($"-Error-\nThere are no languages with ID {intId + 1}.");
                        Console.ReadLine();
                        continue;
                    }
                    newInternalLangNames.Add(internalLangNames[intId]);
                }

                internalLangNames = newInternalLangNames;
            }

            if (answer == "3")
            {
                Console.WriteLine("\nEnter the number of the new main language:");
                int newId = Convert.ToInt32(Console.ReadLine()) - 1;
                if (newId >= internalLangNames.Count)
                {
                    Console.Clear();
                    Console.WriteLine($"-Error-\nThere are no languages with ID {newId + 1}.");
                    Console.ReadLine();
                    continue;
                }

                List<string> newNames = new();
                newNames.Add(internalLangNames[newId]);
                for (int i = 0; i < internalLangNames.Count; i++)
                {
                    if (i == newId)
                    {
                        continue;
                    }
                    newNames.Add(internalLangNames[i]);
                }

                internalLangNames = newNames;
            }
        }

        return internalLangNames;
    }
    
    static List<string> RenameLanguages(List<string> internalLangNames)
    {
        Console.WriteLine("\nDo you want to give these languages alternative names for the spreadsheet?\n\n1 - Yes\n2 - No");
        string answer = Console.ReadLine();
        if (answer == "1")
        {
            Console.Clear();
            List<string> newNames = new();
            for (int i = 0; i < internalLangNames.Count; i++)
            {
                Console.Clear();
                Console.WriteLine($"Enter an alternative name for the {internalLangNames[i]} language (don't type anything to keep the original name):\n");
                var newName = Console.ReadLine();
                if (newName != "")
                {
                    newNames.Add(newName);
                }
                else
                {
                    newNames.Add(internalLangNames[i]);
                }
            }
        
            return newNames;
        }
    
        return internalLangNames;
    }
    
    static MSBP? ParseMSBP(Dictionary<string, string> fileOptions)
    {
        Console.Clear();
        Console.WriteLine("You can provide an MSBP file. It contains all names of control tags, color constants, attributes & more. The info from this file will be used to form human-readable tags and such.\n\nEnter the path to your game's MSBP. If you don't have an MSBP file, type nothing and press Enter.");
        string path;
        if (fileOptions.ContainsKey("msbpPath"))
        {
            path = fileOptions["msbpPath"].Trim('"');
        }
        else
        {
            path = Console.ReadLine().Trim('"');
        }
    
        if (path == "")
        {
            return null;
        }
        else
        {
            return new MSBP(File.OpenRead(path));
        }
    }
    
    static ParsingOptions SetParsingOptions(Dictionary<string, string> fileOptions, MSBP msbp = null)
    {
        Console.Clear();
        ParsingOptions options = new();
        string answer = "";
        
        Console.Clear();
        Console.WriteLine("Do you want to shorten control tags? (eg. <Control.wait msec=250> becomes <wait 250>)\n\n1 - Yes\n2 - No");
        
        if (fileOptions.ContainsKey("shortenTags"))
        {
            options.ShortenTags = Convert.ToBoolean(fileOptions["shortenTags"]);
        }
        else
        {
            answer = Console.ReadLine();
            if (answer == "1")
            {
                options.ShortenTags = true;
            }
        }
        
        Console.Clear();
        Console.WriteLine("Do you want shorten the <PageBreak> tag to just <p>?\n\n1 - Yes\n2 - No");
        
        if (fileOptions.ContainsKey("shortenPagebreak"))
        {
            options.ShortenPagebreak = Convert.ToBoolean(fileOptions["shortenPagebreak"]);
        }
        else
        {
            answer = Console.ReadLine();
            if (answer == "1")
            {
                options.ShortenPagebreak = true;
            }
        }
        
        Console.Clear();
        Console.WriteLine("Do you want to insert a linebreak after every pagebreak tag?\n\n1 - Yes\n2 - No");
        
        if (fileOptions.ContainsKey("addLinebreaksAfterPagebreaks"))
        {
            options.AddLinebreaksAfterPagebreaks = Convert.ToBoolean(fileOptions["addLinebreaksAfterPagebreaks"]);
        }
        else
        {
            answer = Console.ReadLine();
            if (answer == "1")
            {
                options.AddLinebreaksAfterPagebreaks = true;
            }
        }
        
        Console.Clear();
        Console.WriteLine("Do you want to skip Ruby tags (only Japanese)?\n\n1 - Yes\n2 - No");
        
        if (fileOptions.ContainsKey("skipRuby"))
        {
            options.SkipRuby = Convert.ToBoolean(fileOptions["skipRuby"]);
        }
        else
        {
            answer = Console.ReadLine();
            if (answer == "1")
            {
                options.SkipRuby = true;
            }
        }
        
        Console.Clear();
        Console.WriteLine("Do you want to freeze the first two columns in each sheet (labels and main language)?\n\n1 - Freeze both\n2 - Freeze only labels\n3 - Don't freeze");
        
        if (fileOptions.ContainsKey("freezeColumnCount"))
        {
            options.FreezeColumnCount = Convert.ToInt32(fileOptions["freezeColumnCount"]);
        }
        else
        {
            answer = Console.ReadLine();
            if (answer == "2")
            {
                options.FreezeColumnCount = 1;
            }

            if (answer == "3")
            {
                options.FreezeColumnCount = 0;
            }
        }
        
        Console.Clear();
        Console.WriteLine("Do you want to customize the size of the columns? Default is 250.\n\n1 - Yes\n2 - No");
        
        if (fileOptions.ContainsKey("columnSize"))
        {
            options.ColumnSize = Convert.ToInt32(fileOptions["columnSize"]);
        }
        else
        {
            answer = Console.ReadLine();
            if (answer == "1")
            {
                Console.WriteLine("\nEnter a new value:");
                options.ColumnSize = Convert.ToInt32(Console.ReadLine());
            }
        }
        
        Console.Clear();
        Console.WriteLine("Do you want to highlight control tags?\n\n1 - Yes\n2 - No");
        
        if (fileOptions.ContainsKey("highlightTags"))
        {
            options.HighlightTags = Convert.ToBoolean(fileOptions["highlightTags"]);
        }
        else
        {
            answer = Console.ReadLine();
            if (answer == "1")
            {
                options.HighlightTags = true;
            }
        }
        
        AddTranslationLanguages(ref options, fileOptions);

        return options;
    }
    
    static string FindUnnecessaryPathPrefix(string languagePath, string previousPrefix)
    {
        List<string> directories = Directory.GetDirectories(languagePath + previousPrefix).ToList();
        List<string> files = Directory.GetFiles(languagePath + previousPrefix).ToList();

        if (files.Count(x => Path.GetExtension(x) == ".msbt") > 0)
        {
            return previousPrefix;
        }

        List<string> dirsWithMsbt = new List<string>();
        foreach (var dir in directories)
        {
            if (DirectoryHasMsbts(dir))
            {
                dirsWithMsbt.Add(dir);
            }
        }

        if (dirsWithMsbt.Count == 1)
        {
            return FindUnnecessaryPathPrefix(languagePath, previousPrefix + Path.GetFileName(dirsWithMsbt[0]) + '/');
        }
        else
        {
            return previousPrefix;
        }
    }
    
    static List<List<MSBT>> ParseAllMSBTs(string languagesPath, List<string> internalLangNames, ParsingOptions options, MSBP? msbp = null)
    {
        List<List<MSBT>> languages = new();
        foreach (var lang in internalLangNames)
        {
            languages.Add(ParseMSBTFolder(languagesPath + '\\' + lang, "", lang, options, msbp));
        }

        for (int i = 0; i < options.TransLangNames.Count; i++)
        {
            if (options.TransLangPaths[i] != "")
            {
                languages.Add(ParseMSBTFolder(options.TransLangPaths[i], "", options.TransLangNames[i], options, msbp));
            }
            else
            {
                languages.Add(new List<MSBT>());
            }
        }
    
        return languages;
    }
    
    static Spreadsheet LanguagesToSpreadsheet(List<List<MSBT>> langs, List<string> sheetLangNames, ParsingOptions options, Dictionary<string, string> fileOptions, MSBP? msbp = null)
    {
        Console.Clear();
        Console.WriteLine("Enter a name for your spreadsheet:");
        string spreadsheetName;
        if (fileOptions.ContainsKey("spreadsheetName"))
        {
            spreadsheetName = fileOptions["spreadsheetName"];
        }
        else
        {
            spreadsheetName = Console.ReadLine();
        }
        Console.Clear();

        Spreadsheet spreadsheet = new Spreadsheet()
        {
            Properties = new SpreadsheetProperties()
            {
                Title = spreadsheetName
            },
            Sheets = new List<Sheet>()
        };
        
        List<MSBT> baseLang = langs[0];

        for (int j = 0; j < baseLang.Count; j++)
        {
            MSBT msbt = baseLang[j];
            Console.WriteLine($"Creating {msbt.FileName} sheet...");

            ConditionalFormatRule noTransRule = new()
            {
                Ranges = new List<GridRange>()
                {
                    new GridRange()
                    {
                        SheetId = j + 1,
                        StartRowIndex = 1,
                        StartColumnIndex = sheetLangNames.Count + 1
                    }
                },
                BooleanRule = new BooleanRule()
                {
                    Condition = new BooleanCondition()
                    {
                        Type = "TEXT_EQ",
                        Values = new List<ConditionValue>()
                        {
                            new ConditionValue()
                            {
                                UserEnteredValue = "{{no-translation}}"
                            }
                        }
                    },
                    Format = new CellFormat()
                    {
                        BackgroundColorStyle = new ColorStyle()
                        {
                            RgbColor = new Color()
                            {
                                Red = 0.957f,
                                Green = 0.8f,
                                Blue = 0.8f,
                                Alpha = 1
                            }
                        }
                    }
                }
            };

            int columnCount = 1 + langs.Count + options.TransLangNames.Count;
            if (options.TransLangNames.Count == 0 && (msbt.HasATR1 || msbt.HasTSY1))
            {
                columnCount++;
            }
            
            Sheet sheet = new Sheet()
            {
                Properties = new SheetProperties()
                {
                    Title = msbt.FileName,
                    GridProperties = new GridProperties()
                    {
                        RowCount = 1 + msbt.Messages.Count,
                        ColumnCount = columnCount,
                        FrozenRowCount = msbt.Messages.Count != 0 ? 1 : 0,
                        FrozenColumnCount = options.FreezeColumnCount != 1 + langs.Count ? options.FreezeColumnCount : 1
                    },
                    SheetId = j + 1
                },
                Data = new List<GridData>()
                {
                    new GridData()
                    {
                        RowData = new List<RowData>(),
                        ColumnMetadata = new List<DimensionProperties>()
                        {
                            new DimensionProperties()
                            {
                                PixelSize = 200
                            }
                        }
                    }
                },
                ConditionalFormats = new List<ConditionalFormatRule>()
                {
                    options.TransLangNames.Count > 0 ? noTransRule : null
                }
            };

            for (int i = 0; i < sheetLangNames.Count; i++)
            {
                sheet.Data[0].ColumnMetadata.Add(new DimensionProperties(){PixelSize = options.ColumnSize});
            }

            for (int i = 0; i < options.TransLangNames.Count; i++)
            {
                sheet.Data[0].ColumnMetadata.Add(new DimensionProperties(){PixelSize = options.ColumnSize});
                sheet.Data[0].ColumnMetadata.Add(new DimensionProperties(){PixelSize = 100});
            }
                
            RowData headerRow = new()
            {
                Values = new List<CellData>()
            };
            
            headerRow.Values.Add(new CellData()
            {
                UserEnteredValue = new ExtendedValue()
                {
                    StringValue = "Labels"
                },
                UserEnteredFormat = Constants.HEADER_CELL_FORMAT_CENTERED
            });
            
            foreach (var langName in sheetLangNames)
            {
                headerRow.Values.Add(new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = langName
                    },
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT_CENTERED
                });
            }

            for (int i = 0; i < options.TransLangNames.Count; i++)
            {
                string transLangName = options.TransLangNames[i];
                
                headerRow.Values.Add(new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = transLangName
                    },
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT_CENTERED
                });
                headerRow.Values.Add(new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        FormulaValue = $"='#Stats'!{NumToColumnName(4 + i * 2)}3"
                    },
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT_PERCENT_CENTERED
                });
            }

            if (options.TransLangNames.Count == 0 && (msbt.HasATR1 || msbt.HasTSY1))
            {
                headerRow.Values.Add(new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = "Attributes"
                    },
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT_PERCENT_CENTERED
                });
            }
            
            sheet.Data[0].RowData.Add(headerRow);
            
            foreach (var message in msbt.Messages)
            {
                RowData messageRow = new()
                {
                    Values = new List<CellData>()
                };
                messageRow.Values.Add(new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = message.Key.ToString()
                    }
                });
                for (int i = 0; i < sheetLangNames.Count; i++)
                {
                    string messageText = "";
                    List<MSBT> lang = langs[i];
                    MSBT localizedMSBT = lang.FirstOrDefault(x => x.FileName == msbt.FileName);
                    if (localizedMSBT != null)
                    {
                        if (localizedMSBT.Messages.ContainsKey(message.Key))
                        {
                            messageText = localizedMSBT.Messages[message.Key].Text;
                        }
                        else
                        {
                            messageText = "{{no-message}}";
                        }
                    }
                    else
                    {
                        messageText = "{{no-file}}";
                    }
                    messageRow.Values.Add(new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = messageText
                        },
                        TextFormatRuns = options.HighlightTags ? HighlightRuns(messageText) : null
                    });
                }

                CellData attributesCell = new CellData();
                if (msbt.HasATR1 || msbt.HasTSY1)
                {
                    List<string> attributes = new();
                    if (msbt.HasATR1)
                    {
                        attributes.AddRange(message.Value.Attribute.ToStringList(msbp));
                    }

                    if (msbt.HasTSY1)
                    {
                        if (msbp != null && message.Value.StyleId < msbp.Styles.Count && message.Value.StyleId >= 0)
                        {
                            attributes.Add($"Style: {msbp.Styles[message.Value.StyleId].Name}");
                        }
                        else
                        {
                            attributes.Add($"StyleId: {message.Value.StyleId}");
                        }
                    }

                    string attributesCellData = "";
                    foreach (var attr in attributes)
                    {
                        attributesCellData += $"{attr}; ";
                    }
                    attributesCellData = attributesCellData.Trim();

                    attributesCell.UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = attributesCellData
                    };
                    
                    if (options.TransLangNames.Count == 0)
                    {
                        messageRow.Values.Add(attributesCell);
                    }
                }

                for (int i = sheetLangNames.Count; i < langs.Count; i++)
                {
                    string messageText = "";
                    List<MSBT> transLang = langs[i];
                    MSBT localizedMSBT = transLang.FirstOrDefault(x => x.FileName == msbt.FileName);
                    if (localizedMSBT != null)
                    {
                        if (localizedMSBT.Messages.ContainsKey(message.Key))
                        {
                            messageText = localizedMSBT.Messages[message.Key].Text;
                            if (messageText == message.Value.Text)
                            {
                                messageText = "{{no-translation}}";
                            }
                        }
                        else
                        {
                            messageText = "{{no-translation}}";
                        }
                    }
                    else
                    {
                        messageText = "{{no-translation}}";
                    }
                    messageRow.Values.Add(new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = messageText
                        },
                        TextFormatRuns = options.HighlightTags ? HighlightRuns(messageText) : null
                    });
                    
                    messageRow.Values.Add(attributesCell);
                }
                
                sheet.Data[0].RowData.Add(messageRow);
            }
            
            spreadsheet.Sheets.Add(sheet);
        }

        if (msbp != null)
        {
            if (msbp.SourceFileNames.Count > 0)
            {
                AddSourceFilesToSpreadsheet(ref spreadsheet, msbp);
            }
            
            if (msbp.AttributeInfos.Count > 0)
            {
                AddAttributesSheetToSpreadsheet(ref spreadsheet, msbp);
            }

            if (msbp.TagGroups.Count > 0)
            {
                AddTagsSheetToSpreadsheet(ref spreadsheet, msbp);
            }

            if (msbp.Styles.Count > 0)
            {
                AddStylesToSpreadsheet(ref spreadsheet, msbp);
            }

            if (msbp.Colors.Count > 0)
            {
                if (msbp.HasCLB1)
                {
                    AddColorsSheetToSpreadsheet(ref spreadsheet, msbp);
                }
                else
                {
                    AddColorsSheetToSpreadsheet(ref spreadsheet, msbp, false);
                }
            }
        }
        
        AddInternalMSBTDataToSpreadsheet(ref spreadsheet, langs[0]);
        AddSettingsToSpreadsheet(ref spreadsheet, options);
        AddStatsToSpreadsheet(ref spreadsheet, sheetLangNames, options.TransLangNames);

        return spreadsheet;
    }
    
    static void HighlightTtydChanges(ref Spreadsheet switchSpreadsheet, GoogleSheetsManager sheetsManager, string gcnSpreadsheetId)
    {
        Console.Clear();
        Console.WriteLine("Loading the GCN TTYD spreadsheet metadata...");
        Spreadsheet gcnSpreadsheet = sheetsManager.GetSpreadSheet(gcnSpreadsheetId);
        
        List<string> targetSheetNames = new();
        foreach (var sheet in gcnSpreadsheet.Sheets)
        {
            string sheetTitle = sheet.Properties.Title;
            if (!sheetTitle.StartsWith('#'))
            {
                targetSheetNames.Add(sheetTitle);
            }
        }
        
        List<string> requestRanges = new List<string>();
        foreach (string sheetName in targetSheetNames)
        {
            requestRanges.Add($"{sheetName}!A:ZZZ");
        }
        
        Console.WriteLine("Loading game text from the GCN TTYD spreadsheet...");
        BatchGetValuesResponse valueRanges = sheetsManager.GetMultipleValues(gcnSpreadsheetId, requestRanges.ToArray());
        var values = valueRanges.ValueRanges.ToList();
        
        foreach (var sheet in switchSpreadsheet.Sheets)
        {
            if (sheet.Properties.Title.StartsWith('#'))
            {
                continue;
            }
            
            for (int i = 1; i < sheet.Data[0].RowData.Count; i++)
            {
                RowData row = sheet.Data[0].RowData[i];
                string label = row.Values[0].UserEnteredValue.StringValue;

                if (label.EndsWith('%') || label == "Labels")
                {
                    continue;
                }
                
                string engNewMessage = row.Values[1].UserEnteredValue.StringValue;
                string japNewMessage = row.Values[2].UserEnteredValue.StringValue;
                string gerNewMessage = row.Values[6].UserEnteredValue.StringValue;
                string fraNewMessage = row.Values[7].UserEnteredValue.StringValue;
                string itaNewMessage = row.Values[9].UserEnteredValue.StringValue;
                string spaNewMessage = row.Values[10].UserEnteredValue.StringValue;
                
                Console.WriteLine($"Analysing {sheet.Properties.Title}@{label}...");

                var origMessages = FindOrigMessages(label, sheet.Properties.Title, targetSheetNames, values);
                string engOrigMessage = origMessages[0];
                string japOrigMessage = origMessages[1];
                string gerOrigMessage = origMessages[2];
                string fraOrigMessage = origMessages[3];
                string itaOrigMessage = origMessages[4];
                string spaOrigMessage = origMessages[5];

                if (engOrigMessage == "{{not-found}}")
                {
                    row.Values[1].UserEnteredFormat = new CellFormat()
                    {
                        BackgroundColorStyle = new ColorStyle()
                        {
                            RgbColor = new Color()
                            {
                                Red = 0.812f, Green = 0.886f, Blue = 0.953f
                            }
                        }
                    };
                    row.Values[2].UserEnteredFormat = new CellFormat()
                    {
                        BackgroundColorStyle = new ColorStyle()
                        {
                            RgbColor = new Color()
                            {
                                Red = 0.812f, Green = 0.886f, Blue = 0.953f
                            }
                        }
                    };
                    row.Values[6].UserEnteredFormat = new CellFormat()
                    {
                        BackgroundColorStyle = new ColorStyle()
                        {
                            RgbColor = new Color()
                            {
                                Red = 0.812f, Green = 0.886f, Blue = 0.953f
                            }
                        }
                    };
                    row.Values[7].UserEnteredFormat = new CellFormat()
                    {
                        BackgroundColorStyle = new ColorStyle()
                        {
                            RgbColor = new Color()
                            {
                                Red = 0.812f, Green = 0.886f, Blue = 0.953f
                            }
                        }
                    };
                    row.Values[9].UserEnteredFormat = new CellFormat()
                    {
                        BackgroundColorStyle = new ColorStyle()
                        {
                            RgbColor = new Color()
                            {
                                Red = 0.812f, Green = 0.886f, Blue = 0.953f
                            }
                        }
                    };
                    row.Values[10].UserEnteredFormat = new CellFormat()
                    {
                        BackgroundColorStyle = new ColorStyle()
                        {
                            RgbColor = new Color()
                            {
                                Red = 0.812f, Green = 0.886f, Blue = 0.953f
                            }
                        }
                    };
                }
                else
                {
                    string normalizedEngNewMessage = NormalizeMessage(engNewMessage);
                    string normalizedJapNewMessage = NormalizeMessage(japNewMessage);
                    string normalizedGerNewMessage = NormalizeMessage(gerNewMessage);
                    string normalizedFraNewMessage = NormalizeMessage(fraNewMessage);
                    string normalizedItaNewMessage = NormalizeMessage(itaNewMessage);
                    string normalizedSpaNewMessage = NormalizeMessage(spaNewMessage);
                    string normalizedEngOrigMessage = NormalizeMessage(engOrigMessage);
                    string normalizedJapOrigMessage = NormalizeMessage(japOrigMessage);
                    string normalizedGerOrigMessage = NormalizeMessage(gerOrigMessage);
                    string normalizedFraOrigMessage = NormalizeMessage(fraOrigMessage);
                    string normalizedItaOrigMessage = NormalizeMessage(itaOrigMessage);
                    string normalizedSpaOrigMessage = NormalizeMessage(spaOrigMessage);

                    if (normalizedEngNewMessage != normalizedEngOrigMessage)
                    {
                        row.Values[1].UserEnteredFormat = new CellFormat()
                        {
                            BackgroundColorStyle = new ColorStyle()
                            {
                                RgbColor = new Color()
                                {
                                    Red = 1, Green = 0.949f, Blue = 0.8f
                                }
                            }
                        };
                    }
                    if (normalizedJapNewMessage != normalizedJapOrigMessage)
                    {
                        row.Values[2].UserEnteredFormat = new CellFormat()
                        {
                            BackgroundColorStyle = new ColorStyle()
                            {
                                RgbColor = new Color()
                                {
                                    Red = 1, Green = 0.949f, Blue = 0.8f
                                }
                            }
                        };
                    }
                    if (normalizedGerNewMessage != normalizedGerOrigMessage)
                    {
                        row.Values[6].UserEnteredFormat = new CellFormat()
                        {
                            BackgroundColorStyle = new ColorStyle()
                            {
                                RgbColor = new Color()
                                {
                                    Red = 1, Green = 0.949f, Blue = 0.8f
                                }
                            }
                        };
                    }
                    if (normalizedFraNewMessage != normalizedFraOrigMessage)
                    {
                        row.Values[7].UserEnteredFormat = new CellFormat()
                        {
                            BackgroundColorStyle = new ColorStyle()
                            {
                                RgbColor = new Color()
                                {
                                    Red = 1, Green = 0.949f, Blue = 0.8f
                                }
                            }
                        };
                    }
                    if (normalizedItaNewMessage != normalizedItaOrigMessage)
                    {
                        row.Values[9].UserEnteredFormat = new CellFormat()
                        {
                            BackgroundColorStyle = new ColorStyle()
                            {
                                RgbColor = new Color()
                                {
                                    Red = 1, Green = 0.949f, Blue = 0.8f
                                }
                            }
                        };
                    }
                    if (normalizedSpaNewMessage != normalizedSpaOrigMessage)
                    {
                        row.Values[10].UserEnteredFormat = new CellFormat()
                        {
                            BackgroundColorStyle = new ColorStyle()
                            {
                                RgbColor = new Color()
                                {
                                    Red = 1, Green = 0.949f, Blue = 0.8f
                                }
                            }
                        };
                    }
                }
            }
        }
    }
    
    static void AddTranslationLanguages(ref ParsingOptions options, Dictionary<string, string> fileOptions)
    {
        if (fileOptions.ContainsKey("newLangs") && fileOptions.ContainsKey("newLangsPaths"))
        {
            options.TransLangNames.AddRange(fileOptions["newLangs"].Split('|'));
            options.TransLangPaths.AddRange(fileOptions["newLangsPaths"].Split('|'));
            return;
        }
    
        Console.Clear();
        Console.WriteLine("Do you plan on translating this game? If so, type the number of languages you want to translate it into. If not, type 0.");
        int answer = Convert.ToInt32(Console.ReadLine());
        for (int i = 0; i < answer; i++)
        {
            Console.Clear();
            Console.WriteLine($"Type the name of the {i + 1}th language (this is how it would appear in the sheets):");
            string name = Console.ReadLine();
            options.TransLangNames.Add(name);
        
            Console.WriteLine("\nDo you have an unfinished translation for this language? We could incorporate it into the sheets. If so, type the path to the language folder:");
            string path = Console.ReadLine();
            if (path != "")
            {
                ConsoleUtils.CheckDirectory(path);
            }
            options.TransLangPaths.Add(path);
        }
    }
    
    static bool DirectoryHasMsbts(string path)
    {
        List<string> directories = Directory.GetDirectories(path).ToList();
        List<string> files = Directory.GetFiles(path).ToList();

        if (files.Count(x => Path.GetExtension(x) == ".msbt") > 0)
        {
            return true;
        }

        foreach (var dir in directories)
        {
            if (DirectoryHasMsbts(path + '/' + Path.GetFileName(dir)))
            {
                return true;
            }
        }

        return false;
    }
    
    static List<MSBT> ParseMSBTFolder(string folderPath, string fileNamePrefix, string internalLangName, ParsingOptions options, MSBP? msbp = null)
    {
        Console.Clear();
        List<MSBT> msbts = new();
    
        string[] filePaths = Directory.GetFiles(folderPath);
        foreach (var filePath in filePaths)
        {
            if (Path.GetExtension(filePath) != ".msbt")
            {
                continue;
            }
        
            Console.WriteLine($"Parsing {filePath}...");
            MSBT msbt = new(File.OpenRead(filePath), options, fileNamePrefix + Path.GetFileNameWithoutExtension(filePath), internalLangName, msbp);
            msbts.Add(msbt);
        }

        string[] directoryPaths = Directory.GetDirectories(folderPath);
        List<MSBT> msbtsInSubdirectories = new();
        foreach (var directoryPath in directoryPaths)
        {
            msbtsInSubdirectories.AddRange(ParseMSBTFolder(directoryPath, fileNamePrefix + Path.GetFileName(directoryPath) + "/", internalLangName, options, msbp));
        }
        msbts.AddRange(msbtsInSubdirectories);

        return msbts;
    }
    
    static string NumToColumnName(int columnNumber)
    {
        columnNumber++;
        string columnName = "";

        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        } 

        return columnName;
    }
    
    static List<TextFormatRun> HighlightRuns(string message)
    {
        List<int> leftBracketPoses = AllIndexesOf(message, '<');
        List<int> rightBracketPoses = AllIndexesOf(message, '>');

        TextFormat highlightFormat = new TextFormat()
        {
            ForegroundColorStyle = new ColorStyle()
            {
                RgbColor = new Color()
                {
                    Red = 1,
                    Green = 0,
                    Blue = 0,
                    Alpha = 1
                }
            }
        };
    
        TextFormat noHighlightFormat = new TextFormat()
        {
            ForegroundColorStyle = new ColorStyle()
            {
                RgbColor = new Color()
                {
                    Red = 0,
                    Green = 0,
                    Blue = 0,
                    Alpha = 1
                }
            }
        };
    
        List<TextFormatRun> runs = new();
        foreach (var leftBracketPos in leftBracketPoses)
        {
            runs.Add(new TextFormatRun()
            {
                StartIndex = leftBracketPos,
                Format = highlightFormat
            });
        }
    
        foreach (var rightBracketPos in rightBracketPoses)
        {
            runs.Add(new TextFormatRun()
            {
                StartIndex = rightBracketPos + 1,
                Format = noHighlightFormat
            });
        }

        return runs;
    }
    
    static void AddStatsToSpreadsheet(ref Spreadsheet spreadsheet, List<string> origLangNames, List<string> transLangNames)
    {
        Sheet statsSheet = new()
        {
            Properties = new SheetProperties()
            {
                Title = "#Stats",
                SheetId = 0,
                GridProperties = new GridProperties()
                {
                    RowCount = 3 + spreadsheet.Sheets.Where(x => !x.Properties.Title.StartsWith('#')).ToList().Count,
                    ColumnCount = 3 + 2 * transLangNames.Count,
                    FrozenRowCount = 3
                }
            },
            Data = new List<GridData>()
            {
                new GridData()
                {
                    RowData = new List<RowData>()
                }
            },
            ConditionalFormats = new List<ConditionalFormatRule>(),
            Merges = new List<GridRange>()
        };

        if (transLangNames.Count > 0)
        {
            var zeroPercentRule = new ConditionalFormatRule()
            {
                Ranges = new List<GridRange>(),
                BooleanRule = new BooleanRule()
                {
                    Condition = new BooleanCondition()
                    {
                        Type = "NUMBER_EQ",
                        Values = new List<ConditionValue>()
                        {
                            new ConditionValue()
                            {
                                UserEnteredValue = "0%"
                            }
                        }
                    },
                    Format = new CellFormat()
                    {
                        BackgroundColorStyle = new ColorStyle()
                        {
                            RgbColor = new Color()
                            {
                                Red = 0.918f,
                                Green = 0.6f,
                                Blue = 0.6f,
                                Alpha = 1
                            }
                        }
                    }
                }
            };
            
            var hundredPercentRule = new ConditionalFormatRule()
            {
                Ranges = new List<GridRange>(),
                BooleanRule = new BooleanRule()
                {
                    Condition = new BooleanCondition()
                    {
                        Type = "NUMBER_EQ",
                        Values = new List<ConditionValue>()
                        {
                            new ConditionValue()
                            {
                                UserEnteredValue = "100%"
                            }
                        }
                    },
                    Format = new CellFormat()
                    {
                        BackgroundColorStyle = new ColorStyle()
                        {
                            RgbColor = new Color()
                            {
                                Red = 0.851f,
                                Green = 0.918f,
                                Blue = 0.827f,
                                Alpha = 1
                            }
                        }
                    }
                }
            };
        
            for (int i = 0; i < transLangNames.Count; i++)
            {
                GridRange percentRange = new GridRange()
                {
                    SheetId = statsSheet.Properties.SheetId,
                    StartRowIndex = 3,
                    StartColumnIndex = 4 + i * 2,
                    EndColumnIndex = 5 + i * 2
                };
            
                zeroPercentRule.Ranges.Add(percentRange);
                hundredPercentRule.Ranges.Add(percentRange);
            }
        
            statsSheet.ConditionalFormats.Add(zeroPercentRule);
            statsSheet.ConditionalFormats.Add(hundredPercentRule);
        }

        RowData langsHeaderRow = new()
        {
            Values = new List<CellData>()
            {
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = "File Info"
                    },
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT_CENTERED
                },
                new CellData()
                {
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT_CENTERED
                },
                new CellData()
                {
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT_CENTERED
                },
            }
        };
        statsSheet.Merges.Add(new GridRange()
        {
            SheetId = statsSheet.Properties.SheetId,
            StartColumnIndex = 0,
            StartRowIndex = 0,
            EndColumnIndex = 3,
            EndRowIndex = 1
        });

        for (int i = 0; i < transLangNames.Count; i++)
        {
            string langName = transLangNames[i];
            langsHeaderRow.Values.Add(new CellData()
            {
                UserEnteredValue = new ExtendedValue()
                {
                    StringValue = langName
                },
                UserEnteredFormat = Constants.HEADER_CELL_FORMAT_CENTERED_LEFT_BORDER
            });
            langsHeaderRow.Values.Add(new CellData()
            {
                UserEnteredFormat = Constants.HEADER_CELL_FORMAT
            });
            
            statsSheet.Merges.Add(new GridRange()
            {
                SheetId = statsSheet.Properties.SheetId,
                StartRowIndex = 0,
                EndRowIndex = 1,
                StartColumnIndex = 3 + 2 * i,
                EndColumnIndex = 5 + 2 * i
            });
        }
        
        statsSheet.Data[0].RowData.Add(langsHeaderRow);

        RowData headerRow = new()
        {
            Values = new List<CellData>()
            {
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = "Filename"
                    },
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT
                },
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = "Total\nMessages"
                    },
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT
                },
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = "Total\nCharacters"
                    },
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT
                },
            }
        };

        for (int i = 0; i < transLangNames.Count(); i++)
        {
            headerRow.Values.Add(new CellData()
            {
                UserEnteredValue = new ExtendedValue()
                {
                    StringValue = "Untranslated\nCharacters"
                },
                UserEnteredFormat = Constants.HEADER_CELL_FORMAT_LEFT_BORDER
            });
            headerRow.Values.Add(new CellData()
            {
                UserEnteredValue = new ExtendedValue()
                {
                    StringValue = "Done by"
                },
                UserEnteredFormat = Constants.HEADER_CELL_FORMAT
            });
        }
        
        statsSheet.Data[0].RowData.Add(headerRow);
        
        RowData totalRow = new()
        {
            Values = new List<CellData>()
            {
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = "Total:"
                    },
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT
                },
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        FormulaValue = "=SUM(B4:B)"
                    },
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT
                },
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        FormulaValue = "=SUM(C4:C)"
                    },
                    UserEnteredFormat = Constants.HEADER_CELL_FORMAT
                }
            }
        };

        for (int i = 0; i < transLangNames.Count; i++)
        {
            totalRow.Values.Add(new CellData()
            {
                UserEnteredValue = new ExtendedValue()
                {
                    FormulaValue = $"=SUM({NumToColumnName(3 + i * 2)}4:{NumToColumnName(3 + i * 2)})"
                },
                UserEnteredFormat = Constants.HEADER_CELL_FORMAT_LEFT_BORDER
            });
            
            totalRow.Values.Add(new CellData()
            {
                UserEnteredValue = new ExtendedValue()
                {
                    FormulaValue = $"=IF(C3<>0; (C3-{NumToColumnName(3 + i * 2)}3)/C3; 0)"
                },
                UserEnteredFormat = Constants.HEADER_CELL_FORMAT_PERCENT
            });
        }
        
        statsSheet.Data[0].RowData.Add(totalRow);
        
        int sheetNum = 0;
        foreach (var sheet in spreadsheet.Sheets)
        {
            if (sheet.Properties.Title.StartsWith('#'))
            {
                continue;
            }

            sheetNum++;
            
            RowData row = new()
            {
                Values = new List<CellData>()
                {
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            FormulaValue = $"=HYPERLINK(\"#gid={sheet.Properties.SheetId}\";\"{sheet.Properties.Title}\")"
                        },
                    },
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            FormulaValue = $"=COUNTA('{sheet.Properties.Title}'!A2:A)"
                        }
                    },
                    new CellData()
                    {
                        UserEnteredValue = sheet.Data[0].RowData.Count > 1 ? new ExtendedValue()
                        {
                            FormulaValue = $"=SUMPRODUCT(LEN('{sheet.Properties.Title}'!B2:B))"
                        } : new ExtendedValue()
                        {
                            NumberValue = 0
                        }
                    }
                }
            };

            for (int i = 0; i < transLangNames.Count(); i++)
            {
                string transColumnName = NumToColumnName(1 + origLangNames.Count + i * 2);

                row.Values.Add(new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        FormulaValue =
                            $"=C{sheetNum + 3}-SUM(ARRAYFORMULA(IF('{sheet.Properties.Title}'!{transColumnName}2:{transColumnName} <> \"{{{{no-translation}}}}\", LEN('{sheet.Properties.Title}'!B2:B), 0)))"
                    },
                    UserEnteredFormat = new CellFormat()
                    {
                        Borders = new Borders()
                        {
                            Left = new Border()
                            {
                                Style = "SOLID"
                            }
                        }
                    }
                });
                
                row.Values.Add(new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        FormulaValue = $"=IF(C{sheetNum + 3}<>0; (C{sheetNum + 3}-{NumToColumnName(3 + i * 2)}{sheetNum + 3})/C{sheetNum + 3}; 0)"
                    },
                    UserEnteredFormat = new CellFormat()
                    {
                        NumberFormat = new NumberFormat()
                        {
                            Type = "PERCENT",
                            Pattern = "0.00%"
                        },
                        BackgroundColorStyle = new ColorStyle()
                        {
                            RgbColor = new Color()
                            {
                                Red = 0.957f, Blue = 0.8f, Green = 0.8f, Alpha = 1
                            } 
                        }
                    }
                });
            }
            
            statsSheet.Data[0].RowData.Add(row);
        }
        
        spreadsheet.Sheets.Insert(0, statsSheet);
    }

    static void AddInternalMSBTDataToSpreadsheet(ref Spreadsheet spreadsheet, List<MSBT> baseLang)
    {
        Sheet dataSheet = new Sheet()
        {
            Properties = new SheetProperties()
            {
                Title = "#InternalData",
                GridProperties = new GridProperties()
                {
                    RowCount = 1 + baseLang.Count,
                    ColumnCount = baseLang[0].HasATO1 ? 6 : 5,
                    FrozenRowCount = 1
                }
            },
            Data = new List<GridData>()
            {
                new GridData()
                {
                    RowData = new List<RowData>()
                }
            }
        };
        
        RowData headerRow = new()
        {
            Values = StringListToCellData(
                new List<string>(){"Filename", "Slot Count", "Version", "Byte Order", "Encoding"},
                Constants.HEADER_CELL_FORMAT
            )
        };

        if (baseLang[0].HasATO1)
        {
            headerRow.Values.Add(new CellData()
            {
                UserEnteredValue = new ExtendedValue()
                {
                    StringValue = "ATO1 Section"
                },
                UserEnteredFormat = Constants.HEADER_CELL_FORMAT
            });
        }
        
        dataSheet.Data[0].RowData.Add(headerRow);

        foreach (var msbt in baseLang)
        {
            RowData row = new()
            {
                Values = new List<CellData>()
                {
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = msbt.FileName
                        },
                    },
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            NumberValue = msbt.LabelSlotCount
                        }
                    },
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            NumberValue = msbt.Header.Version
                        }
                    },
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = msbt.Header.Endianness == Endianness.BigEndian ? "Big Endian" : "Little Endian"
                        }
                    },
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = msbt.Header.EncodingType.ToString()
                        }
                    }
                }
            };

            if (baseLang[0].HasATO1)
            {
                string ato1String = "";
                foreach (var num in msbt.ATO1Numbers)
                {
                    ato1String += $"{num}, ";
                }
                ato1String = ato1String.TrimEnd(' ');
                ato1String = ato1String.TrimEnd(',');
                
                row.Values.Add(new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = ato1String
                    },
                    UserEnteredFormat = new CellFormat()
                    {
                        WrapStrategy = "WRAP"
                    }
                });
            }
            
            dataSheet.Data[0].RowData.Add(row);
        }
        
        spreadsheet.Sheets.Insert(0, dataSheet);
    }

    static void AddColorsSheetToSpreadsheet(ref Spreadsheet spreadsheet, MSBP msbp, bool hasNames = true)
    {
        Sheet colorSheet = new Sheet()
        {
            Properties = new SheetProperties()
            {
                Title = "#BaseColors",
                GridProperties = new GridProperties()
                {
                    ColumnCount = hasNames ? 2 : 1,
                    RowCount = msbp.Colors.Count
                }
            },
            Data = new List<GridData>()
            {
                new GridData()
                {
                    RowData = new List<RowData>()
                }
            }
        };

        foreach (var color in msbp.Colors)
        {
            if (hasNames)
            {
                colorSheet.Data[0].RowData.Add(new RowData()
                {
                    Values = new List<CellData>()
                    {
                        new CellData()
                        {
                            UserEnteredValue = new ExtendedValue()
                            {
                                StringValue = color.Key
                            },
                            UserEnteredFormat = CellColorFromMsbpColor(color.Value)
                        },
                        new CellData()
                        {
                            UserEnteredValue = new ExtendedValue()
                            {
                                StringValue = GeneralUtils.ColorToString(color.Value)
                            }
                        }
                    }
                });
            }
            else
            {
                colorSheet.Data[0].RowData.Add(new RowData()
                {
                    Values = new List<CellData>()
                    {
                        new CellData()
                        {
                            UserEnteredValue = new ExtendedValue()
                            {
                                StringValue = GeneralUtils.ColorToString(color.Value)
                            },
                            UserEnteredFormat = CellColorFromMsbpColor(color.Value)
                        }
                    }
                });
            }
        }
        
        spreadsheet.Sheets.Insert(0, colorSheet);
    }

    static CellFormat CellColorFromMsbpColor(System.Drawing.Color color)
    {
        float floatR = color.R / 255f;
        float floatG = color.G / 255f;
        float floatB = color.B / 255f;
        float floatA = color.A / 255f;

        return new CellFormat()
        {
            BackgroundColorStyle = new ColorStyle()
            {
                RgbColor = new Color()
                {
                    Red = 1 - floatA + floatA * floatR,
                    Green = 1 - floatA + floatA * floatG,
                    Blue = 1 - floatA + floatA * floatB
                }
            },
            TextFormat = new TextFormat()
            {
                ForegroundColorStyle = new ColorStyle()
                {
                    RgbColor = color.GetBrightness() < 0.5f
                        ? new Color()
                        {
                            Red = 1, Green = 1, Blue = 1
                        }
                        : new Color()
                        {
                            Red = 0, Green = 0, Blue = 0
                        }
                }
            }
        };
    }

    static CellFormat CellColorFromMsbpColorId(int index, MSBP msbp)
    {
        if (MsbpHasColor(index, msbp))
        {
            return CellColorFromMsbpColor(GetColorFromId(index, msbp));
        }

        return new CellFormat();
    }

    static void AddSettingsToSpreadsheet(ref Spreadsheet spreadsheet, ParsingOptions options)
    {
        Sheet sheet = new Sheet()
        {
            Properties = new SheetProperties()
            {
                Title = "#Settings",
                GridProperties = new GridProperties()
                {
                    ColumnCount = 2,
                    FrozenRowCount = 1
                }
            },
            Data = new List<GridData>()
            {
                new GridData()
                {
                    RowData = new List<RowData>()
                }
            }
        };
        
        RowData headerRow = new()
        {
            Values = StringListToCellData(
                new List<string>(){"Setting", "Value"},
                Constants.HEADER_CELL_FORMAT
            )
        };
        
        sheet.Data[0].RowData.Add(headerRow);
        
        sheet.Data[0].RowData.Add(new RowData()
        {
            Values = new List<CellData>()
            {
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = "Add linebreaks after pagebreaks"
                    },
                    UserEnteredFormat = new CellFormat()
                    {
                        WrapStrategy = "WRAP",
                        TextFormat = new TextFormat()
                        {
                            Bold = true
                        }
                    }
                },
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        BoolValue = options.AddLinebreaksAfterPagebreaks
                    }
                }
            }
        });

        if (options.UnnecessaryPathPrefix != "")
        {
            sheet.Data[0].RowData.Add(new RowData()
            {
                Values = new List<CellData>()
                {
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = "Common MSBT path prefix"
                        },
                        UserEnteredFormat = new CellFormat()
                        {
                            WrapStrategy = "WRAP",
                            TextFormat = new TextFormat()
                            {
                                Bold = true
                            }
                        }
                    },
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = options.UnnecessaryPathPrefix
                        }
                    }
                }
            });
        }
        
        sheet.Data[0].RowData.Add(new RowData()
        {
            Values = new List<CellData>()
            {
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = "Color identification"
                    },
                    UserEnteredFormat = new CellFormat()
                    {
                        WrapStrategy = "WRAP",
                        TextFormat = new TextFormat()
                        {
                            Bold = true
                        }
                    }
                },
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = options.ColorIdentification
                    }
                }
            }
        });
        
        sheet.Properties.GridProperties.RowCount = sheet.Data[0].RowData.Count;
        spreadsheet.Sheets.Insert(0, sheet);
    }

    static void AddStylesToSpreadsheet(ref Spreadsheet spreadsheet, MSBP msbp)
    {
        Sheet sheet = new Sheet()
        {
            Properties = new SheetProperties()
            {
                Title = "#Styles",
                GridProperties = new GridProperties()
                {
                    ColumnCount = 5,
                    RowCount = msbp.Styles.Count + 1,
                    FrozenRowCount = 1
                }
            },
            Data = new List<GridData>()
            {
                new GridData()
                {
                    RowData = new List<RowData>()
                }
            }
        };
        
        RowData headerRow = new()
        {
            Values = StringListToCellData(
                new List<string>(){"Name", "Region Width", "Line Count", "Font Index", "Base Color"},
                Constants.HEADER_CELL_FORMAT
            )
        };
        
        sheet.Data[0].RowData.Add(headerRow);

        foreach (var style in msbp.Styles)
        {
            RowData row = new RowData()
            {
                Values = new List<CellData>()
                {
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = style.Name
                        }
                    },
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            NumberValue = style.RegionWidth
                        }
                    },
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            NumberValue = style.LineCount
                        }
                    },
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            NumberValue = style.FontId
                        }
                    },
                    new CellData()
                    {
                        UserEnteredValue = MsbpHasColor(style.BaseColorId, msbp) ? new ExtendedValue()
                        {
                            StringValue = GeneralUtils.GetColorNameFromId(style.BaseColorId, msbp)
                        } : new ExtendedValue()
                        {
                            NumberValue = style.BaseColorId
                        },
                        UserEnteredFormat = CellColorFromMsbpColorId(style.BaseColorId, msbp)
                    },
                }
            };
            
            sheet.Data[0].RowData.Add(row);
        }
        
        spreadsheet.Sheets.Insert(0, sheet);
    }

    static void AddTagsSheetToSpreadsheet(ref Spreadsheet spreadsheet, MSBP msbp)
    {
        msbp.TagGroups[0].Tags[4].Name = "PageBreak";
        
        Sheet sheet = new Sheet()
        {
            Properties = new SheetProperties()
            {
                Title = "#Tags",
                GridProperties = new GridProperties()
                {
                    ColumnCount = 5,
                    FrozenRowCount = 1
                },
                SheetId = 1000000000
            },
            Data = new List<GridData>()
            {
                new GridData()
                {
                    RowData = new List<RowData>()
                }
            },
            Merges = new List<GridRange>()
        };
        
        RowData headerRow = new()
        {
            Values = StringListToCellData(
                new List<string>(){"Group", "Tag", "Parameter", "Param. Type", "List Items"},
                Constants.HEADER_CELL_FORMAT
            )
        };
        
        sheet.Data[0].RowData.Add(headerRow);

        foreach (var group in msbp.TagGroups)
        {
            bool groupWritten = false;
            foreach (var tag in group.Tags)
            {
                bool tagWritten = false;
                foreach (var param in tag.Parameters)
                {
                    RowData row = new RowData()
                    {
                        Values = new List<CellData>()
                        {
                            new CellData()
                            {
                                UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = !groupWritten ? $"{group.Id}. {group.Name}" : ""
                                }
                            },
                            new CellData()
                            {
                                UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = !tagWritten ? tag.Name : ""
                                }
                            },
                            new CellData()
                            {
                                UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = param.Name
                                }
                            },
                            new CellData()
                            {
                                UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = ParamTypeToString(param.Type)
                                }
                            }
                        }
                    };

                    groupWritten = true;
                    tagWritten = true;

                    if (param.Type == ParamType.List)
                    {
                        if (param.List.Count > 0)
                        {
                            row.Values.Add(new CellData()
                            {
                                UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = param.List[0]
                                }
                            });
                            
                            sheet.Data[0].RowData.Add(row);
                            
                            for (int i = 1; i < param.List.Count; i++)
                            {
                                sheet.Data[0].RowData.Add(new RowData()
                                {
                                    Values = new List<CellData>()
                                    {
                                        new CellData()
                                        {
                                            UserEnteredValue = new ExtendedValue()
                                            {
                                                StringValue = ""
                                            }
                                        },
                                        new CellData()
                                        {
                                            UserEnteredValue = new ExtendedValue()
                                            {
                                                StringValue = ""
                                            }
                                        },
                                        new CellData()
                                        {
                                            UserEnteredValue = new ExtendedValue()
                                            {
                                                StringValue = ""
                                            }
                                        },
                                        new CellData()
                                        {
                                            UserEnteredValue = new ExtendedValue()
                                            {
                                                StringValue = ""
                                            }
                                        },
                                        new CellData()
                                        {
                                            UserEnteredValue = new ExtendedValue()
                                            {
                                                StringValue = param.List[i]
                                            }
                                        }
                                    }
                                });
                            }
                        }
                    }
                    else
                    {
                        row.Values.Add(new CellData()
                            {
                                UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = ""
                                }
                            }
                        );
                        sheet.Data[0].RowData.Add(row);
                    }
                }

                if (tag.Parameters.Count == 0)
                {
                    RowData row = new RowData()
                    {
                        Values = new List<CellData>()
                        {
                            new CellData()
                            {
                                UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = !groupWritten ? $"{group.Id}. {group.Name}" : ""
                                }
                            },
                            new CellData()
                            {
                                UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = tag.Name
                                }
                            },
                            new CellData()
                            {
                                UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = ""
                                }
                            },
                            new CellData()
                            {
                                UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = ""
                                }
                            },
                            new CellData()
                            {
                                UserEnteredValue = new ExtendedValue()
                                {
                                    StringValue = ""
                                }
                            }
                        }
                    };
                    
                    groupWritten = true;
                    sheet.Data[0].RowData.Add(row);
                }
            }

            if (group.Tags.Count == 0)
            {
                RowData row = new RowData()
                {
                    Values = new List<CellData>()
                    {
                        new CellData()
                        {
                            UserEnteredValue = new ExtendedValue()
                            {
                                StringValue = $"{group.Id}. {group.Name}"
                            }
                        },
                        new CellData()
                        {
                            UserEnteredValue = new ExtendedValue()
                            {
                                StringValue = ""
                            }
                        },
                        new CellData()
                        {
                            UserEnteredValue = new ExtendedValue()
                            {
                                StringValue = ""
                            }
                        },
                        new CellData()
                        {
                            UserEnteredValue = new ExtendedValue()
                            {
                                StringValue = ""
                            }
                        },
                        new CellData()
                        {
                            UserEnteredValue = new ExtendedValue()
                            {
                                StringValue = ""
                            }
                        }
                    }
                };

                sheet.Data[0].RowData.Add(row);
            }
        }

        sheet.Properties.GridProperties.RowCount = sheet.Data[0].RowData.Count;
        MergeEmptyVertically(ref sheet, 0);
        MergeEmptyVertically(ref sheet, 1);
        MergeEmptyParamsVertically(ref sheet);
        CenterCellsVertically(ref sheet);
        
        spreadsheet.Sheets.Insert(0, sheet);
    }

    static void AddSourceFilesToSpreadsheet(ref Spreadsheet spreadsheet, MSBP msbp)
    {
        Sheet sheet = new Sheet()
        {
            Properties = new SheetProperties()
            {
                Title = "#SourceFiles",
                GridProperties = new GridProperties()
                {
                    RowCount = msbp.SourceFileNames.Count,
                    ColumnCount = 1
                }
            },
            Data = new List<GridData>()
            {
                new GridData()
                {
                    RowData = new List<RowData>()
                }
            }
        };

        foreach (var filename in msbp.SourceFileNames)
        {
            sheet.Data[0].RowData.Add(new RowData()
            {
                Values = new List<CellData>()
                {
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = filename
                        }
                    }
                }
            });
        }
        
        spreadsheet.Sheets.Insert(0, sheet);
    }

    static void AddAttributesSheetToSpreadsheet(ref Spreadsheet spreadsheet, MSBP msbp)
    {
        Sheet sheet = new Sheet()
        {
            Properties = new SheetProperties()
            {
                Title = "#Attributes",
                GridProperties = new GridProperties()
                {
                    ColumnCount = 3,
                    FrozenRowCount = 1
                },
                SheetId = 1000000001
            },
            Data = new List<GridData>()
            {
                new GridData()
                {
                    RowData = new List<RowData>()
                }
            },
            Merges = new List<GridRange>()
        };
        
        RowData headerRow = new()
        {
            Values = StringListToCellData(
                new List<string>(){"Name", "Type", "List Items"},
                Constants.HEADER_CELL_FORMAT
            )
        };
        
        sheet.Data[0].RowData.Add(headerRow);

        foreach (var attr in msbp.AttributeInfos)
        {
            RowData row = new RowData()
            {
                Values = new List<CellData>()
                {
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = attr.Name
                        }
                    },
                    new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = ParamTypeToString(attr.Type)
                        }
                    },
                }
            };
            
            if (attr.Type == ParamType.List)
            {
                if (attr.List.Count > 0)
                {
                    row.Values.Add(new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = attr.List[0]
                        }
                    });
                    
                    sheet.Data[0].RowData.Add(row);
                    
                    for (int i = 1; i < attr.List.Count; i++)
                    {
                        sheet.Data[0].RowData.Add(new RowData()
                        {
                            Values = new List<CellData>()
                            {
                                new CellData()
                                {
                                    UserEnteredValue = new ExtendedValue()
                                    {
                                        StringValue = ""
                                    }
                                },
                                new CellData()
                                {
                                    UserEnteredValue = new ExtendedValue()
                                    {
                                        StringValue = ""
                                    }
                                },
                                new CellData()
                                {
                                    UserEnteredValue = new ExtendedValue()
                                    {
                                        StringValue = attr.List[i]
                                    }
                                }
                            }
                        });
                    }
                }
            }
            else
            {
                row.Values.Add(new CellData()
                    {
                        UserEnteredValue = new ExtendedValue()
                        {
                            StringValue = ""
                        }
                    }
                );
                sheet.Data[0].RowData.Add(row);
            }
        }
        
        MergeEmptyVertically(ref sheet, 0);
        MergeEmptyVertically(ref sheet, 1);
        CenterCellsVertically(ref sheet);
        
        sheet.Properties.GridProperties.RowCount = sheet.Data[0].RowData.Count;
        spreadsheet.Sheets.Insert(0, sheet);
    }

    static void MergeEmptyVertically(ref Sheet sheet, int columnIndex)
    {
        for (int i = 1; i < sheet.Data[0].RowData.Count; i++)
        {
            int height = 1;
            for (int j = i + 1; j < sheet.Data[0].RowData.Count; j++)
            {
                RowData lowerRow = sheet.Data[0].RowData[j];
                string lowerValue = lowerRow.Values[columnIndex].UserEnteredValue.StringValue;
                string lowerLeftValue = lowerRow.Values[0].UserEnteredValue.StringValue;
                bool isLastRow = j + 1 == sheet.Data[0].RowData.Count;
                if (lowerValue != "" || lowerLeftValue != "" || isLastRow && lowerValue == "" && lowerLeftValue == "")
                {
                    if (isLastRow && lowerValue == "" && lowerLeftValue == "")
                    {
                        height++;
                    }
                    
                    sheet.Merges.Add(new GridRange()
                    {
                        SheetId = sheet.Properties.SheetId,
                        StartColumnIndex = columnIndex,
                        EndColumnIndex = columnIndex + 1,
                        StartRowIndex = i,
                        EndRowIndex = i + height
                    });

                    i += height - 1;
                    break;
                }
                
                height++;
            }
        }
    }

    static void MergeEmptyParamsVertically(ref Sheet sheet)
    {
        for (int i = 1; i < sheet.Data[0].RowData.Count; i++)
        {
            string paramType = sheet.Data[0].RowData[i].Values[3].UserEnteredValue.StringValue;
            if (paramType != "One of:")
            {
                continue;
            }
            
            int height = 1;
            for (int j = i + 1; j < sheet.Data[0].RowData.Count; j++)
            {
                RowData lowerRow = sheet.Data[0].RowData[j];
                bool isLastRow = j + 1 == sheet.Data[0].RowData.Count;
                string lowerGroup = lowerRow.Values[0].UserEnteredValue.StringValue;
                string lowerTag = lowerRow.Values[1].UserEnteredValue.StringValue;
                string lowerParam = lowerRow.Values[2].UserEnteredValue.StringValue;
                string lowerParamType = lowerRow.Values[3].UserEnteredValue.StringValue;
                if (lowerParam != "" || lowerParamType != "" || lowerGroup != "" || lowerTag != "" ||
                    isLastRow && lowerParam == "" && lowerParamType == "" && lowerTag == "" && lowerGroup == "")
                {
                    if (isLastRow && lowerParam == "" && lowerParamType == "" && lowerTag == "" && lowerGroup == "")
                    {
                        height++;
                    }
                    
                    sheet.Merges.Add(new GridRange()
                    {
                        SheetId = sheet.Properties.SheetId,
                        StartColumnIndex = 2,
                        EndColumnIndex = 3,
                        StartRowIndex = i,
                        EndRowIndex = i + height
                    });
                    sheet.Merges.Add(new GridRange()
                    {
                        SheetId = sheet.Properties.SheetId,
                        StartColumnIndex = 3,
                        EndColumnIndex = 4,
                        StartRowIndex = i,
                        EndRowIndex = i + height
                    });

                    i += height - 1;
                    break;
                }
                else
                {
                    height++;
                }
            }
        }
    }

    static void CenterCellsVertically(ref Sheet sheet)
    {
        foreach (var row in sheet.Data[0].RowData)
        {
            foreach (var cell in row.Values)
            {
                if (cell.UserEnteredFormat != null)
                {
                    cell.UserEnteredFormat.VerticalAlignment = "MIDDLE";
                }
                else
                {
                    cell.UserEnteredFormat = new CellFormat()
                    {
                        VerticalAlignment = "MIDDLE"
                    };
                }
            }
        }
    }

    static string ParamTypeToString(ParamType type)
    {
        switch (type)
        {
            case ParamType.UInt8:
                return "UInt8";
            case ParamType.UInt16:
                return "UInt16";
            case ParamType.UInt32:
                return "UInt32";
            case ParamType.Int8:
                return "Int8";
            case ParamType.Int16:
                return "Int16";
            case ParamType.Int32:
                return "Int32";
            case ParamType.Float:
                return "Float";
            case ParamType.Double:
                return "Double";
            case ParamType.String:
                return "String";
            case ParamType.List:
                return "One of:";
            default:
                return "Unknown type";
        }
    }
    
    static System.Drawing.Color GetColorFromId(int index, MSBP msbp)
    {
        uint i = 0;
        foreach (var color in msbp.Colors)
        {
            if (i == index)
            {
                return color.Value;
            }

            i++;
        }

        return System.Drawing.Color.White;
    }

    static bool MsbpHasColor(int index, MSBP msbp)
    {
        return index >= 0 && index < msbp.Colors.Count;
    }
    
    static List<CellData> StringListToCellData(List<string> texts, CellFormat format = null, int? repeatCount = 1)
    {
        List<CellData> cellDatas = new();

        for (int i = 0; i < repeatCount; i++)
        {
            foreach (var text in texts)
            {
                CellData cellData = new()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = text
                    },
                    UserEnteredFormat = format
                };
            
                cellDatas.Add(cellData);
            }
        }

        return cellDatas;
    }
    
    static List<int> AllIndexesOf(string str, char searchChar)
    {
        var foundIndexes = new List<int>();
        for (int i = str.IndexOf(searchChar); i > -1; i = str.IndexOf(searchChar, i + 1))
        {
            foundIndexes.Add(i);
        }

        return foundIndexes;
    }
    
    static string NormalizeMessage(string message)
    {
        message = message.Replace("\n", "");
        message = message.Replace(" ", "");
        message = Regex.Replace(message, @"<[^>]+>", "");

        return message;
    }

    static List<string> FindOrigMessages(string label, string switchSheetName, List<string> gcnSheetNames, List<ValueRange> valueRanges)
    {
        for (int i = 0; i < gcnSheetNames.Count; i++)
        {
            if (gcnSheetNames[i].StartsWith(switchSheetName))
            {
                foreach (var row in valueRanges[i].Values)
                {
                    if ((string)row[0] == label)
                    {
                        return new List<string>()
                        {
                            (string)row[1],
                            (string)row[2],
                            (string)row[3],
                            (string)row[4],
                            (string)row[5],
                            (string)row[6],
                        };
                    }
                }
            }
        }
        
        for (int i = 0; i < gcnSheetNames.Count; i++)
        {
            if (!gcnSheetNames[i].StartsWith(switchSheetName))
            {
                foreach (var row in valueRanges[i].Values)
                {
                    if ((string)row[0] == label)
                    {
                        return new List<string>()
                        {
                            (string)row[1],
                            (string)row[2],
                            (string)row[3],
                            (string)row[4],
                            (string)row[5],
                            (string)row[6],
                        };
                    }
                }
            }
        }
        
        return new List<string>()
        {
            "{{not-found}}",
            "{{not-found}}",
            "{{not-found}}",
            "{{not-found}}",
            "{{not-found}}",
            "{{not-found}}",
        };
    }
}
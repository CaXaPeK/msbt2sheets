using System.Diagnostics;
using System.Net.Mime;
using System.Text;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;using Google.Apis.Sheets.v4.Data;
using Msbt2Sheets.Lib.Formats;
using Msbt2Sheets.Sheets;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

string credPath = Environment.CurrentDirectory + "/data/credentials.txt";
CheckFile(credPath);

string googleClientId = "";
string googleClientSecret = "";
string[] scopes = { SheetsService.Scope.Spreadsheets };

foreach (string s in File.ReadAllLines(credPath))
{
    if (s.StartsWith("clientId="))
        googleClientId = s.Split('=')[1];
    else if (s.StartsWith("clientSecret="))
        googleClientSecret = s.Split('=')[1];
}

UserCredential credential = GoogleAuth.Login(googleClientId, googleClientSecret, scopes);
GoogleSheetsManager sheetsManager = new GoogleSheetsManager(credential);
//Spreadsheet spreadsheet = sheetsManager.GetSpreadSheet("1JkQJUkN3zgfn1FczgZepGnbcyfiWsibKZN79xorPlSo");

//Console.WriteLine("asd");

while (true)
{
    Console.Clear();
    Console.WriteLine("Welcome to Msbt2Sheets! What do you want to create?\n\n1 - Spreadsheet\n2 - MSBT");
    var mode = Console.ReadLine();
    switch (mode)
    {
        case "1":
            MsbtToSpreadsheet(sheetsManager);
            Exit();
            break;
        case "2":
            //do other thing
            break;
        default:
            break;
    }
}

static void MsbtToSpreadsheet(GoogleSheetsManager sheetsManager)
{
    Console.Clear();
    Console.WriteLine("Enter the path to the language folder (folder with EU_English, EU_French, etc.):");
    string languagesPath = Console.ReadLine().Trim('"');
    CheckDirectory(languagesPath);

    List<string> internalLangNames = ReorderLanguages(languagesPath);
    List<string> sheetLangNames = RenameLanguages(internalLangNames);
    
    

    MSBP msbp = ParseMSBP();
    ParsingOptions options = SetParsingOptions(msbp);
    
    List<List<MSBT>> languages = ParseAllMSBTs(languagesPath, internalLangNames, options, msbp);

    Spreadsheet spreadsheet = LanguagesToSpreadsheet(languages, sheetLangNames, options);
    
    Console.WriteLine("Uploading your spreadsheet...");
    var newSheet = sheetsManager.CreateSpreadsheet(spreadsheet);
    
    Console.WriteLine($"Congratulations! Your spreadsheet is ready:\n{newSheet.SpreadsheetUrl}\n\nHave fun!");
    Console.ReadLine();
}

static void AddTranslationLanguages(ref ParsingOptions options)
{
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
            CheckDirectory(path);
        }
        options.TransLangPaths.Add(path);
    }
}

static List<string> ReorderLanguages(string langPath)
{
    List<string> langPaths = Directory.GetDirectories(langPath).ToList();
    List<string> internalLangNames = new();
    foreach (var path in langPaths)
    {
        internalLangNames.Add(Path.GetFileName(path));
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

static ParsingOptions SetParsingOptions(MSBP msbp = null)
{
    Console.Clear();
    ParsingOptions options = new();
    string answer = "";
    
    if (msbp != null)
    {
        Console.Clear();
        Console.WriteLine("Do you want to shorten control tags? (eg. <Control.wait msec=250> becomes <wait 250>)\n\n1 - Yes\n2 - No");
        answer = Console.ReadLine();
        if (answer == "1")
        {
            options.ShortenTags = true;
        }
    }
    
    Console.Clear();
    Console.WriteLine("Do you want shorten the <PageBreak> tag to just <p>?\n\n1 - Yes\n2 - No");
    answer = Console.ReadLine();
    if (answer == "1")
    {
        options.ShortenPagebreak = true;
    }
    
    Console.Clear();
    Console.WriteLine("Do you want to insert a linebreak after every pagebreak tag?\n\n1 - Yes\n2 - No");
    answer = Console.ReadLine();
    if (answer == "1")
    {
        options.AddLinebreaksAfterPagebreaks = true;
    }
    
    Console.Clear();
    Console.WriteLine("Do you want to skip Ruby tags (only Japanese)?\n\n1 - Yes\n2 - No");
    answer = Console.ReadLine();
    if (answer == "1")
    {
        options.SkipRuby = true;
    }
    
    Console.Clear();
    Console.WriteLine("Do you want to freeze the first two columns in each sheet (labels and main language)?\n\n1 - Freeze both\n2 - Freeze only labels\n3 - Don't freeze");
    answer = Console.ReadLine();
    if (answer == "2")
    {
        options.FreezeColumnCount = 1;
    }

    if (answer == "3")
    {
        options.FreezeColumnCount = 0;
    }
    
    Console.Clear();
    Console.WriteLine("Do you want to customize the size of the columns? Default is 250.\n\n1 - Yes\n2 - No");
    answer = Console.ReadLine();
    if (answer == "1")
    {
        Console.WriteLine("\nEnter a new value:");
        options.ColumnSize = Convert.ToInt32(Console.ReadLine());
    }
    
    Console.Clear();
    Console.WriteLine("Do you want to highlight control tags?\n\n1 - Yes\n2 - No");
    answer = Console.ReadLine();
    if (answer == "1")
    {
        options.HighlightTags = true;
    }
    
    AddTranslationLanguages(ref options);

    return options;
}

static MSBP? ParseMSBP()
{
    Console.Clear();
    Console.WriteLine("You can provide an MSBP file. It contains all names of control tags, color constants, attributes & more. The info from this file will be used to form human-readable tags and such.\n\nEnter the path to your game's MSBP. If you don't have an MSBP file, type nothing and press Enter.");
    var path = Console.ReadLine().Trim('"');
    if (path == "")
    {
        return null;
    }
    else
    {
        return new MSBP(File.OpenRead(path));
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
        languages.Add(ParseMSBTFolder(options.TransLangPaths[i], "", options.TransLangNames[i], options, msbp));
    }
    
    return languages;
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
        msbtsInSubdirectories.AddRange(ParseMSBTFolder(directoryPath, Path.GetFileName(directoryPath), internalLangName, options, msbp));
    }
    msbts.AddRange(msbtsInSubdirectories);

    return msbts;
}

static Spreadsheet LanguagesToSpreadsheet(List<List<MSBT>> langs, List<string> sheetLangNames, ParsingOptions options)
{
    Console.Clear();
    Console.WriteLine("Enter a name for your spreadsheet:");
    string spreadsheetName = Console.ReadLine();
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
                    SheetId = j,
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
        
        Sheet sheet = new Sheet()
        {
            Properties = new SheetProperties()
            {
                Title = msbt.FileName,
                GridProperties = new GridProperties()
                {
                    RowCount = 1 + msbt.Messages.Count,
                    ColumnCount = 1 + langs.Count,
                    FrozenRowCount = 1,
                    FrozenColumnCount = options.FreezeColumnCount != 1 + langs.Count ? options.FreezeColumnCount : 1
                },
                SheetId = j
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

        for (int i = 0; i < langs.Count; i++)
        {
            sheet.Data[0].ColumnMetadata.Add(new DimensionProperties(){PixelSize = options.ColumnSize});
        }
            
        RowData headerRow = new()
        {
            Values = new List<CellData>()
        };
        headerRow.Values.Add(new CellData()
        {
            UserEnteredValue = new ExtendedValue()
            {
                StringValue = "Label"
            }
        });
        foreach (var langName in sheetLangNames)
        {
            headerRow.Values.Add(new CellData()
            {
                UserEnteredValue = new ExtendedValue()
                {
                    StringValue = langName
                }
            });
        }

        foreach (var transLangName in options.TransLangNames)
        {
            headerRow.Values.Add(new CellData()
            {
                UserEnteredValue = new ExtendedValue()
                {
                    StringValue = transLangName
                }
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
            }
            
            sheet.Data[0].RowData.Add(messageRow);
        }
        
        spreadsheet.Sheets.Add(sheet);
    }

    return spreadsheet;
}

static void AddStatsToSpreadsheet(ref Spreadsheet spreadsheet)
{
    Sheet statsSheet = new()
    {
        Properties = new SheetProperties()
        {
            Title = "#Stats",
            GridProperties = new GridProperties()
            {
                RowCount = 1 + spreadsheet.Sheets.Count,
                ColumnCount = 4,
                FrozenColumnCount = 1
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
            new List<string>(){"Filename", "Done by", "Untranslated Characters", "Total Characters"},
            new CellFormat()
            {
                TextFormat = new TextFormat()
                {
                    Bold = true
                },
                BackgroundColorStyle = new ColorStyle()
                {
                    RgbColor = new Color()
                    {
                        Red = 0.937f, Blue = 0.937f, Green = 0.937f, Alpha = 1
                    } 
                }
            }
        )
    };

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
                UserEnteredFormat = new CellFormat()
                {
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
                    StringValue = "=C2/D2*100"
                }
            },
            new CellData()
            {
                UserEnteredValue = new ExtendedValue()
                {
                    StringValue = "=SUM(C3:C)"
                }
            },
            new CellData()
            {
                UserEnteredValue = new ExtendedValue()
                {
                    StringValue = "=SUM(D3:D)"
                }
            }
        }
    };

    statsSheet.Data[0].RowData.Add(headerRow);

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
                        StringValue = sheet.Properties.Title
                    },
                    Hyperlink = $"#gid={sheet.Properties.SheetId}"
                },
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = $"=C{sheetNum + 3}/D{sheetNum + 3}*100"
                    }
                },
                new CellData()
                {
                    UserEnteredValue = new ExtendedValue()
                    {
                        StringValue = $"={sheet.Properties.Title}!SUMPRODUCT(LEN(B2:B))"
                    }
                }
            }
        };
        
        
    }
}

static List<CellData> StringListToCellData(List<string> texts, CellFormat format = null)
{
    List<CellData> cellDatas = new();

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

    return cellDatas;
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

static List<int> AllIndexesOf(string str, char searchChar)
{
    var foundIndexes = new List<int>();
    for (int i = str.IndexOf(searchChar); i > -1; i = str.IndexOf(searchChar, i + 1))
    {
        foundIndexes.Add(i);
    }

    return foundIndexes;
}

static void Error(string err)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine(err);
    Console.ResetColor();
    Console.WriteLine();
    Exit();
}

static void CheckFile(string path)
{
    if (!File.Exists(path))
        Error("File not found: " + path);
}

static void CheckDirectory(string path)
{
    if (!Directory.Exists(path))
        Error("Folder not found: " + path);
}

static void Exit()
{
    Console.WriteLine("Press any key to exit...");
    Console.ReadKey();
    Environment.Exit(0);
}
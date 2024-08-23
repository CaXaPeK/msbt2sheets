using System.Text;
using Google.Apis.Http;
using Google.Apis.Sheets.v4.Data;
using Msbt2Sheets.Lib.Formats;
using Msbt2Sheets.Lib.Formats.FileComponents;
using Msbt2Sheets.Lib.Utils;
using Color = System.Drawing.Color;

namespace Msbt2Sheets.Sheets;

public class SheetsToMsbt
{
    public static void Create(GoogleSheetsManager sheetsManager, Dictionary<string, string> fileOptions)
    {
        ParsingOptions options = new();
        SetOptionsFromFile(options, fileOptions);
        
        Console.Clear();
        ConsoleUtils.WriteLineColored("Enter your spreadsheet's ID.\n(It's in the link: https://docs.google.com/spreadsheets/d/|1pRFVKt4fNnWHKf8kIpSk0qmu7u-EdHEUGwkTP9Kzq3A|/edit)", ConsoleColor.Cyan);
        string spreadsheetId = options.SpreadsheetId != "" ? options.SpreadsheetId : Console.ReadLine();
        
        Console.Clear();
        Console.WriteLine("Loading metadata from the spreadsheet...");
        Spreadsheet spreadsheet = sheetsManager.GetSpreadSheet(spreadsheetId);

        if (options.SheetNames.Count > 0)
        {
            Spreadsheet newSpreadsheet = new()
            {
                Sheets = new List<Sheet>()
            };
            foreach (var sheet in spreadsheet.Sheets)
            {
                if (sheet.Properties.Title.StartsWith('#') || options.SheetNames.Contains(sheet.Properties.Title))
                {
                    newSpreadsheet.Sheets.Add(sheet);
                }
            }

            spreadsheet = newSpreadsheet;
        }
        
        List<string> requestRanges = new List<string>();
        foreach (Sheet sheet in spreadsheet.Sheets)
        {
            requestRanges.Add($"{sheet.Properties.Title}!A:ZZZ");
        }
        
        Console.WriteLine("Loading cell data from the spreadsheet...");
        BatchGetValuesResponse valueRanges = sheetsManager.GetMultipleValues(spreadsheetId, requestRanges.ToArray());
        var spreadsheetValues = valueRanges.ValueRanges.ToList();
        var sheets = ValueRangesToStringLists(spreadsheetValues);

        ObtainOptions(spreadsheet, sheets, options);

        MSBP msbp = ObtainMsbp(spreadsheet, sheets);

        if (options.RecreateSources)
        {
            RecreateSources(msbp.SourceFileNames, options.OutputPath);
            ConsoleUtils.Exit();
        }

        AskLanguageNames(spreadsheet, sheets, options);

        List<List<MSBT>> langs = ObtainMsbts(spreadsheet, sheets, options, msbp);

        AskOutputPath(options);
        
        SaveMsbts(langs, msbp, options);
        
        ConsoleUtils.Exit();
    }

    static void RecreateSources(List<string> sources, string outputPath)
    {
        foreach (var source in sources)
        {
            Console.WriteLine($"Saving {source}");
            string filePath = $"{outputPath}\\{source}";
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));
            File.WriteAllBytes(filePath, new byte[]{});
        }
        
        Console.WriteLine($"\nDone! Saved at {outputPath}");
    }

    static List<List<List<string>>> ValueRangesToStringLists(List<ValueRange> valueRanges)
    {
        List<List<List<string>>> sheets = new();
        foreach (var valueRange in valueRanges)
        {
            var rows = valueRange.Values;
            var sheet = new List<List<string>>();
            foreach (var row in rows)
            {
                List<string> stringRow = new();
                foreach (var cell in row)
                {
                    stringRow.Add((string)cell);
                }
                sheet.Add(stringRow);
            }
            sheets.Add(sheet);
        }

        return sheets;
    }

    static void ObtainOptions(Spreadsheet spreadsheet, List<List<List<string>>> sheets, ParsingOptions options)
    {
        if (!options.NoSettingsSheet)
        {
            Dictionary<string, string> langNames = new();
            int sheetId = spreadsheet.Sheets.ToList().FindIndex(x => x.Properties.Title == "#Settings");
            foreach (var row in sheets[sheetId])
            {
                if (row[0] == "Add linebreaks after pagebreaks")
                {
                    options.AddLinebreaksAfterPagebreaks = row[1] == "TRUE";
                }
                if (row[0] == "Color identification")
                {
                    options.ColorIdentification = row[1];
                }
                if (row[0].EndsWith(" internal name"))
                {
                    string uiName = row[0][..(row[0].Length - " internal name".Length)];
                    string internalName = row[1];
                    langNames.Add(uiName, internalName);
                }
            }
            
            int messageSheetId = spreadsheet.Sheets.ToList().FindIndex(x => !x.Properties.Title.StartsWith('#'));
            List<string> headerRow = sheets[messageSheetId][0];
                
            for (int i = 0; i < headerRow.Count; i++)
            {
                string name = headerRow[i];
                if (name == "Labels" || name == "Attributes" || name.EndsWith('%'))
                {
                    continue;
                }
                options.InternalLangNames.Add(name);
            }

            if (langNames.Count != 0)
            {
                foreach (var entry in langNames)
                {
                    options.InternalLangNames[options.InternalLangNames.FindIndex(x => x == entry.Key)] = entry.Value;
                }
            }
        }
    }

    static void SetOptionsFromFile(ParsingOptions options, Dictionary<string, string> fileOptions)
    {
        foreach (var option in fileOptions)
        {
            switch (option.Key)
            {
                case "spreadsheetId":
                    options.SpreadsheetId = option.Value;
                    break;
                case "outputPath":
                    options.OutputPath = option.Value;
                    break;
                case "noStatsSheet":
                    options.NoStatsSheet = Convert.ToBoolean(option.Value);
                    break;
                case "noSettingsSheet":
                    options.NoSettingsSheet = Convert.ToBoolean(option.Value);
                    break;
                case "noInternalDataSheet":
                    options.NoInternalDataSheet = Convert.ToBoolean(option.Value);
                    break;
                case "addLinebreaksAfterPagebreaks":
                    options.AddLinebreaksAfterPagebreaks = Convert.ToBoolean(option.Value);
                    break;
                case "colorIdentification":
                    options.ColorIdentification = option.Value;
                    break;
                case "skipLangIfNotTranslated":
                    options.SkipLangIfNotTranslated = Convert.ToBoolean(option.Value);
                    break;
                case "extendedHeader":
                    options.ExtendedHeader = Convert.ToBoolean(option.Value);
                    break;
                case "sheetNames":
                    options.SheetNames = option.Value.Split('|').ToList();
                    break;
                case "customFileNames":
                    options.CustomFileNames = option.Value.Split('|').ToList();
                    break;
                case "globalVersion":
                    options.GlobalVersion = Convert.ToByte(option.Value);
                    break;
                case "globalEndianness":
                    options.GlobalEndianness = option.Value == "Little Endian" ? Endianness.LittleEndian : Endianness.BigEndian;
                    break;
                case "globalEncoding":
                    Enum.TryParse(option.Value, out EncodingType encoding);
                    options.GlobalEncodingType = encoding;
                    break;
                case "globalAto1":
                    string[] ato1sStr = option.Value.Split(", ");
                    foreach (var ato1Str in ato1sStr)
                    {
                        options.GlobalAto1.Add(Convert.ToInt32(ato1Str));
                    }
                    break;
                case "slotCounts":
                    string[] slotCountsStr = option.Value.Split('|');
                    foreach (var slotCountStr in slotCountsStr)
                    {
                        options.SlotCounts.Add(Convert.ToUInt32(slotCountStr));
                    }
                    break;
                case "uiLangs":
                    options.UiLangNames = option.Value.Split('|').ToList();
                    break;
                case "outputLangs":
                    options.OutputLangNames = option.Value.Split('|').ToList();
                    break;
                case "noTranslationSymbol":
                    options.NoTranslationSymbol = option.Value;
                    break;
                case "noMessageSymbol":
                    options.NoMessageSymbol = option.Value;
                    break;
                case "noFileSymbol":
                    options.NoFileSymbol = option.Value;
                    break;
                case "mainLangColumnId":
                    options.MainLangColumnId = Convert.ToInt32(option.Value);
                    break;
                case "recreateSources":
                    options.RecreateSources = Convert.ToBoolean(option.Value);
                    break;
            }
        }

        if (options.UiLangNames.Count > 0 && options.OutputLangNames.Count == 0)
        {
            options.OutputLangNames = options.UiLangNames;
        }
    }

    static MSBP ObtainMsbp(Spreadsheet spreadsheet, List<List<List<string>>> sheets)
    {
        MSBP msbp = new();

        ObtainMsbpColors(spreadsheet, sheets, msbp);
        ObtainMsbpStyles(spreadsheet, sheets, msbp);
        ObtainMsbpAttributes(spreadsheet, sheets, msbp);
        ObtainMsbpTags(spreadsheet, sheets, msbp);
        ObtainMsbpSources(spreadsheet, sheets, msbp);

        if (msbp.TagGroups.Count == 0)
        {
            msbp.TagGroups.Add(MSBP.BaseMSBP.TagGroups[0]);
        }

        return msbp;
    }

    static void ObtainMsbpColors(Spreadsheet spreadsheet, List<List<List<string>>> sheets, MSBP msbp)
    {
        int colorSheetId = spreadsheet.Sheets.ToList().FindIndex(x => x.Properties.Title == "#BaseColors");
        if (colorSheetId != -1)
        {
            msbp.HasCLR1 = true;
            Sheet sheet = spreadsheet.Sheets[colorSheetId];
            if (sheet.Properties.GridProperties.ColumnCount == 2)
            {
                msbp.HasCLB1 = true;

                foreach (var row in sheets[colorSheetId])
                {
                    msbp.Colors.Add(row[0], ColorStringToColor(row[1]));
                }
            }
            else
            {
                int counter = 0;
                foreach (var row in sheets[colorSheetId])
                {
                    msbp.Colors.Add(counter.ToString(), ColorStringToColor(row[0]));
                    counter++;
                }
            }
        }
    }
    
    static void ObtainMsbpSources(Spreadsheet spreadsheet, List<List<List<string>>> sheets, MSBP msbp)
    {
        int sourceSheetId = spreadsheet.Sheets.ToList().FindIndex(x => x.Properties.Title == "#SourceFiles");
        if (sourceSheetId != -1)
        {
            msbp.HasCTI1 = true;
            foreach (var row in sheets[sourceSheetId])
            {
                msbp.SourceFileNames.Add(row[0]);
            }
        }
    }

    static Color ColorStringToColor(string colorStr)
    {
        byte[] colorBytes = Convert.FromHexString(colorStr[1..colorStr.Length]);
        return Color.FromArgb(colorBytes[3], colorBytes[0], colorBytes[1], colorBytes[2]);
    }

    static void ObtainMsbpStyles(Spreadsheet spreadsheet, List<List<List<string>>> sheets, MSBP msbp)
    {
        int styleSheetId = spreadsheet.Sheets.ToList().FindIndex(x => x.Properties.Title == "#Styles");
        if (styleSheetId != -1)
        {
            msbp.HasSYL3 = true;
            var rows = sheets[styleSheetId];
            var headerRow = rows[0];
            if (headerRow[0] == "Name")
            {
                msbp.HasSLB1 = true;
            }

            for (int i = 1; i < rows.Count; i++)
            {
                var row = rows[i];

                int colorId = msbp.HasCLR1 ? GetColorIdFromName(row[4], msbp) : Convert.ToInt32(row[4]);

                msbp.Styles.Add(new Style(row[0], Convert.ToInt32(row[1]), Convert.ToInt32(row[2]), Convert.ToInt32(row[3]), colorId));
            }
        }
    }

    static int GetColorIdFromName(string name, MSBP msbp)
    {
        int id = 0;
        foreach (var color in msbp.Colors)
        {
            if (color.Key == name)
            {
                return id;
            }

            id++;
        }

        return Convert.ToInt32(name);
    }

    static void ObtainMsbpAttributes(Spreadsheet spreadsheet, List<List<List<string>>> sheets, MSBP msbp)
    {
        int attrSheetId = spreadsheet.Sheets.ToList().FindIndex(x => x.Properties.Title == "#Attributes");
        if (attrSheetId != -1)
        {
            msbp.HasATI2 = true;
            var rows = sheets[attrSheetId];
            var headerRow = rows[0];
            if (headerRow[0] == "Name")
            {
                msbp.HasALB1 = true;
            }
            if (headerRow[2] == "List Items")
            {
                msbp.HasALB1 = true;
            }

            uint offset = 0;
            for (int i = 1; i < rows.Count; i++)
            {
                var row = rows[i];
                var name = row[0];
                var type = StringToParamType(row[1]);
                var list = new List<string>();
                if (type == ParamType.List)
                {
                    list.Add(row[2]);
                    i++;
                    while (rows[i][0] == "")
                    {
                        list.Add(rows[i][2]);
                        i++;
                        if (i >= rows.Count) break;
                    }

                    i--;
                }
                
                msbp.AttributeInfos.Add(new AttributeInfo(name, type, offset, list));
                offset += (uint)GeneralUtils.GetTypeSize(type);
            }
        }
    }
    
    static ParamType StringToParamType(string type)
    {
        switch (type)
        {
            case "UInt8":
                return ParamType.UInt8;
            case "UInt16":
                return ParamType.UInt16;
            case "UInt32":
                return ParamType.UInt32;
            case "Int8":
                return ParamType.Int8;
            case "Int16":
                return ParamType.Int16;
            case "Int32":
                return ParamType.Int32;
            case "Float":
                return ParamType.Float;
            case "Double":
                return ParamType.Double;
            case "String":
                return ParamType.String;
            case "One of:":
                return ParamType.List;
            default:
                throw new InvalidDataException($"Can't convert an unknown parameter type {type}.");
        }
    }

    static void ObtainMsbpTags(Spreadsheet spreadsheet, List<List<List<string>>> sheets, MSBP msbp)
    {
        int tagSheetId = spreadsheet.Sheets.ToList().FindIndex(x => x.Properties.Title == "#Tags");
        if (tagSheetId != -1)
        {
            msbp.HasTGG2 = true;
            msbp.HasTAG2 = true;
            msbp.HasTGP2 = true;
            msbp.HasTGL2 = true;
            var rows = sheets[tagSheetId];

            for (int i = 1; i < rows.Count; i++)
            {
                var row = rows[i];
                FillWithEmpty(row, 5);
                
                if (row[0] != "")
                {
                    var groupId = Convert.ToUInt16(row[0].Split('.')[0]);
                    var groupName = row[0].Substring(row[0].IndexOf(' ') + 1);
                    msbp.TagGroups.Add(new TagGroup()
                    {
                        Id = groupId,
                        Name = groupName
                    });
                }

                if (row[1] != "")
                {
                    msbp.TagGroups.Last().Tags.Add(new TagType()
                    {
                        Name = row[1]
                    });
                }

                if (row[2] != "")
                {
                    msbp.TagGroups.Last().Tags.Last().Parameters.Add(new TagParameter()
                    {
                        Name = row[2],
                        Type = StringToParamType(row[3])
                    });
                }
                
                if (row[4] != "")
                {
                    msbp.TagGroups.Last().Tags.Last().Parameters.Last().List.Add(row[4]);
                }
            }
        }
    }

    static void FillWithEmpty(List<string> strings, int itemCount)
    {
        for (int i = strings.Count; i < itemCount; i++)
        {
            strings.Add("");
        }
    }

    static void AskLanguageNames(Spreadsheet spreadsheet, List<List<List<string>>> sheets, ParsingOptions options)
    {
        if (options.UiLangNames.Count > 0)
        {
            return;
        }
        
        int msbtSheetId = spreadsheet.Sheets.ToList().FindIndex(x => !x.Properties.Title.StartsWith('#'));
        if (msbtSheetId == -1)
        {
            ConsoleUtils.Exit("Your spreadsheet doesn't contain sheets with messages.");
        }

        List<string> headerRow = sheets[msbtSheetId][0];
        List<string> languageNames = new();
        foreach (var cell in headerRow)
        {
            if (cell != "Labels" && cell != "Attributes" && !cell.EndsWith('%'))
            {
                languageNames.Add(cell);
            }
        }

        List<string> wantedLanguageNames = new();
        List<string> wantedInternalLanguageNames = new();
        while (true)
        {
            Console.Clear();

            Console.WriteLine("The spreadsheet contains the following languages:");
            for (int i = 0; i < languageNames.Count; i++)
            {
                string langName = languageNames[i];
                Console.WriteLine($"{i + 1}: {langName}");
            }

            Console.WriteLine("\nEnter which languages do you want to export (eg. 1 2 9). Type nothing if you want all.");
            var answer = Console.ReadLine();
            if (answer == "")
            {
                wantedLanguageNames = languageNames;
                break;
            }
            var newIds = answer.Trim(' ').Split(' ');

            foreach (var id in newIds)
            {
                int intId = Convert.ToInt32(id) - 1;
                if (intId >= languageNames.Count)
                {
                    Console.Clear();
                    Console.WriteLine($"-Error-\nThere are no languages with ID {intId + 1}.");
                    Console.ReadLine();
                    continue;
                }

                wantedLanguageNames.Add(languageNames[intId]);
                wantedInternalLanguageNames.Add(options.InternalLangNames[intId]);
            }

            break;
        }

        options.UiLangNames = wantedLanguageNames;
        if (options.UiLangNames.Count > 0 && options.OutputLangNames.Count == 0)
        {
            options.OutputLangNames = wantedInternalLanguageNames;
        }
    }

    static List<List<MSBT>> ObtainMsbts(Spreadsheet spreadsheet, List<List<List<string>>> sheets,
        ParsingOptions options, MSBP msbp)
    {
        Console.Clear();
        
        List<List<MSBT>> langs = new();
        for (int i = 0; i < options.UiLangNames.Count; i++)
        {
            langs.Add(new());
        }
        
        var msbtSheetIds = GetMsbtSheetIds(spreadsheet, options);
        for (int a = 0; a < msbtSheetIds.Count; a++)
        {
            int msbtSheetId = msbtSheetIds[a];
            
            var sheet = spreadsheet.Sheets[msbtSheetId];
            var sheetName = sheet.Properties.Title;
            var fileName = options.CustomFileNames.Count == 0 ? sheetName : options.CustomFileNames[a];
            var sheetGrid = sheets[msbtSheetId];
            var headerRow = sheetGrid[0];

            uint slotCount;
            byte version;
            Endianness byteOrder;
            EncodingType encoding;
            int msbpAttributeCount = msbp != null ? msbp.AttributeInfos.Count : 0;
            
            if (!options.NoInternalDataSheet)
            {
                var internalSheetId = spreadsheet.Sheets.ToList().FindIndex(x => x.Properties.Title == "#InternalData");
                if (internalSheetId == -1)
                {
                    ConsoleUtils.Exit("Your spreadsheet doesn't contain an #InternalData sheet.");
                }
            
                var internalDataRowId = sheets[internalSheetId].FindIndex(x => x[0] == sheetName);
                var internalDataRow = sheets[internalSheetId][internalDataRowId];
                slotCount = Convert.ToUInt32(internalDataRow[1]);
                version = Convert.ToByte(internalDataRow[2]);
                byteOrder = internalDataRow[3] == "Little Endian" ? Endianness.LittleEndian : Endianness.BigEndian;
                Enum.TryParse(internalDataRow[4], out EncodingType encodingTemp);
                encoding = encodingTemp;
            }
            else
            {
                if (options.SlotCounts.Count != msbtSheetIds.Count)
                {
                    ConsoleUtils.Exit("Amount of SlotCounts in the preset doesn't correspond with the amount of processed sheets.");
                }

                slotCount = options.SlotCounts[a];
                version = options.GlobalVersion;
                byteOrder = options.GlobalEndianness;
                encoding = options.GlobalEncodingType;
            }
            
            for (int j = 0; j < options.UiLangNames.Count; j++)
            {
                var langName = options.UiLangNames[j];
                Console.WriteLine($"Obtaining {fileName}.msbt ({langName})...");
                var langColumnId = headerRow.IndexOf(langName);
                
                var hasAtr1 = false;
                var hasTsy1 = false;
                var attributeColumnId = -1;
                
                if (headerRow.Contains("Attributes"))
                {
                    attributeColumnId = headerRow.IndexOf("Attributes");
                }

                if (headerRow.Count(x => x.EndsWith('%')) > 0)
                {
                    if (headerRow[langColumnId + 1].EndsWith('%'))
                    {
                        attributeColumnId = langColumnId + 1;
                    }
                    else
                    {
                        attributeColumnId = headerRow.FindIndex(x => x.EndsWith('%'));
                    }
                }

                /*if (sheet.Properties.GridProperties.RowCount == 1 && attributeColumnId != -1)
                {
                    hasAtr1 = true;
                    hasTsy1 = true;
                }*/

                bool noFileFlag = false;
                var messages = new Dictionary<object, Message>();
                int startRowId = options.ExtendedHeader ? 2 : 1;
                bool hasTranslatedMessages = false;
                for (int i = startRowId; i < sheetGrid.Count; i++)
                {
                    var messageRow = sheetGrid[i];

                    FillWithEmpty(messageRow, headerRow.Count);
                    
                    var label = messageRow[0];
                    var text = messageRow[langColumnId];
                    
                    if (text == options.NoMessageSymbol)
                    {
                        continue;
                    }
                    if (text == options.NoFileSymbol)
                    {
                        noFileFlag = true;
                        break;
                    }
                    if (text == options.NoTranslationSymbol)
                    {
                        text = messageRow[options.MainLangColumnId];
                    }
                    else
                    {
                        hasTranslatedMessages = true;
                    }
                    
                    var attributeObjectDictionary = new Dictionary<string, object>();
                    var styleId = -1;
                    
                    if (attributeColumnId != -1)
                    {
                        var attributeCell = messageRow[attributeColumnId];
                        var attributeDict = CellToDictionary(attributeCell);
                        
                        if (attributeDict.ContainsKey("Style"))
                        {
                            hasTsy1 = true;
                            var styleName = attributeDict["Style"];
                            var style = msbp.Styles.FirstOrDefault(x => x.Name == styleName);
                            if (style == null)
                            {
                                styleId = Convert.ToInt32(styleName);
                            }
                            else
                            {
                                styleId = msbp.Styles.IndexOf(style);
                            }

                            attributeDict.Remove("Style");
                        }
                        
                        if (attributeDict.ContainsKey("StyleId"))
                        {
                            hasTsy1 = true;
                            var styleName = attributeDict["StyleId"];
                            styleId = Convert.ToInt32(styleName);

                            attributeDict.Remove("StyleId");
                        }

                        int unnamedAttrCount = attributeDict.Count(x => x.Key.StartsWith("Attribute_"));
                        if (unnamedAttrCount > 0)
                        {
                            msbpAttributeCount = unnamedAttrCount;
                        }

                        if (attributeDict.Count > 0)
                        {
                            hasAtr1 = true;

                            attributeObjectDictionary = StringDictToObjectDict(attributeDict,
                                msbp != null ? msbp.AttributeInfos : new());
                        }
                    }
                    
                    messages.Add(label, new Message
                    {
                        Text = text,
                        StyleId = styleId,
                        Attributes = attributeObjectDictionary
                    });
                }

                if (options.SkipLangIfNotTranslated && !hasTranslatedMessages)
                {
                    continue;
                }
                
                langs[options.UiLangNames.IndexOf(langName)].Add(new MSBT
                {
                    FileName = fileName,
                    Language = options.OutputLangNames[j],
                    Header = new Header
                    {
                        FileType = FileType.MSBT,
                        Version = version,
                        Endianness = byteOrder,
                        EncodingType = encoding
                    },
                    LabelSlotCount = slotCount,
                    HasLBL1 = true,
                    HasATR1 = hasAtr1,
                    HasTSY1 = hasTsy1,
                    HasATO1 = hasAtr1,
                    Messages = messages,
                    MsbpAttributeCount = msbpAttributeCount
                });
            }
        }

        return langs;
    }

    static Dictionary<string, object> StringDictToObjectDict(Dictionary<string, string> strDict,
        List<AttributeInfo> attrInfos)
    {
        Dictionary<string, object> dict = new();

        foreach (var strDictEntry in strDict)
        {
            object value = null;
            
            if (strDictEntry.Key.StartsWith("Attribute_") || strDictEntry.Key == "Attributes")
            {
                value = Convert.FromHexString(strDictEntry.Value.Replace("-", ""));
            }
            
            if (strDictEntry.Key.StartsWith("StringAttributes"))
            {
                List<string> strAttrs = strDictEntry.Value.Split(", ").ToList();
                for (int i = 0; i < strAttrs.Count; i++)
                {
                    strAttrs[i] = GeneralUtils.UnquoteString(strAttrs[i]);
                }

                value = strAttrs;
            }
            
            AttributeInfo attrInfo = attrInfos.FirstOrDefault(x => x.Name == strDictEntry.Key);
            if (attrInfo != null)
            {
                switch (attrInfo.Type)
                {
                    case ParamType.UInt8:
                        value = Convert.ToByte(strDictEntry.Value);
                        break;
                    case ParamType.UInt16:
                        value = Convert.ToUInt16(strDictEntry.Value);
                        break;
                    case ParamType.UInt32:
                        value = Convert.ToUInt32(strDictEntry.Value);
                        break;
                    case ParamType.Int8:
                        value = Convert.ToSByte(strDictEntry.Value);
                        break;
                    case ParamType.Int16:
                        value = Convert.ToInt16(strDictEntry.Value);
                        break;
                    case ParamType.Int32:
                        value = Convert.ToInt32(strDictEntry.Value);
                        break;
                    case ParamType.Float:
                        value = Convert.ToSingle(strDictEntry.Value);
                        break;
                    case ParamType.Double:
                        value = Convert.ToDouble(strDictEntry.Value);
                        break;
                    case ParamType.String:
                        value = strDictEntry.Value;
                        break;
                    case ParamType.List:
                        value = strDictEntry.Value;
                        break;
                    default:
                        throw new InvalidDataException(
                            $"Attribute {attrInfo.Name} has an invalid type of {(byte) attrInfo.Type}");
                }
            }
            
            dict.Add(strDictEntry.Key, value);
        }
        
        return dict;
    }

    static List<int> GetMsbtSheetIds(Spreadsheet spreadsheet, ParsingOptions options)
    {
        List<int> ids = new();
        int counter = 0;
        if (options.SheetNames.Count == 0)
        {
            foreach (var sheet in spreadsheet.Sheets)
            {
                if (!sheet.Properties.Title.StartsWith('#'))
                {
                    ids.Add(counter);
                }

                counter++;
            }
        }
        else
        {
            foreach (var sheet in spreadsheet.Sheets)
            {
                if (options.SheetNames.Contains(sheet.Properties.Title))
                {
                    ids.Add(counter);
                }

                counter++;
            }
        }

        return ids;
    }

    static List<int> CommaSpaceNumbersToList(string cell)
    {
        return cell.Split(", ")
            .Select(numberString => Convert.ToInt32(numberString))
            .ToList();
    }

    static Dictionary<string, string> CellToDictionary(string cell)
    {
        Dictionary<string, string> dict = new();

        while (cell.Length != 0)
        {
            string key = cell[..cell.IndexOf(':')];
            cell = cell[(key.Length + 2)..];

            string value;

            if (cell[0] == '"')
            {
                if (cell.Count(x => x == '"') == 1)
                {
                    throw new InvalidDataException("Can't parse attributes: no closing quote");
                }

                value = cell[..(FindFirstIndexWhereNotAfter(cell, '\\', '"', 1) + 1)];
                cell = cell[value.Length..];
                /*value = value[1..(value.Length - 1)];
                value = value.Replace("\\\"", "\"");
                value = value.Replace("\\\\", "\\");*/
                
                if (cell.Count(x => x == ';') == 0)
                {
                    if (cell.Length == 0)
                    {
                        throw new InvalidDataException("Can't parse attributes: no semi-colon at the end");
                    }
                    else
                    {
                        throw new InvalidDataException("Can't parse attributes: no semi-colon after a string attribute");
                    }
                }
            }
            else
            {
                if (cell.Count(x => x == ';') == 0)
                {
                    throw new InvalidDataException("Can't parse attributes: no semi-colon at the end");
                }
                
                value = cell[..cell.IndexOf(';')];
                cell = cell[value.Length..];
            }
            
            cell = cell[1..]; //remove ; at the start
            if (cell.Length != 0) cell = cell[1..]; //remove space at the start
            
            dict.Add(key, value);
        }

        return dict;
    }

    static int FindFirstIndexWhereNotAfter(string str, char first, char second, int startIndex = 0)
    {
        for (int i = startIndex; i < str.Length; i++)
        {
            if (str[i] == second)
            {
                if (i == 0 || str[i - 1] != first)
                {
                    return i;
                }
            }
        }

        return -1;
    }

    static void AskOutputPath(ParsingOptions options)
    {
        if (options.OutputPath == "")
        {
            Console.Clear();
            Console.WriteLine("Enter the path for outputting MSBT files:");
            var path = Console.ReadLine();

            options.OutputPath = path;
        }
    }

    static void SaveMsbts(List<List<MSBT>> langs, MSBP msbp, ParsingOptions options)
    {
        Console.Clear();

        foreach (var lang in langs)
        {
            if (options.SkipLangIfNotTranslated && lang.Count == 0)
            {
                continue;
            }
            foreach (var msbt in lang)
            {
                Console.WriteLine($"Saving {msbt.Language}\\{options.UnnecessaryPathPrefix}{msbt.FileName}.msbt");
                string filePath = $"{options.OutputPath}\\{msbt.Language}\\{options.UnnecessaryPathPrefix}{msbt.FileName}.msbt";
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                File.WriteAllBytes(filePath, msbt.Compile(options, msbp));
            }
        }
        
        Console.WriteLine($"\nDone! Saved at {options.OutputPath}");
        ConsoleUtils.Exit();
    }
}
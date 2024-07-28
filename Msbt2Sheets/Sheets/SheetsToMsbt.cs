using System.Text;
using Google.Apis.Sheets.v4.Data;
using Msbt2Sheets.Lib.Formats;
using Msbt2Sheets.Lib.Formats.FileComponents;
using Msbt2Sheets.Lib.Utils;
using Color = System.Drawing.Color;

namespace Msbt2Sheets.Sheets;

public class SheetsToMsbt
{
    public static void Create(GoogleSheetsManager sheetsManager)
    {
        Console.Clear();
        ConsoleUtils.WriteLineColored("Enter your spreadsheet's ID.\n(It's in the link: https://docs.google.com/spreadsheets/d/|1pRFVKt4fNnWHKf8kIpSk0qmu7u-EdHEUGwkTP9Kzq3A|/edit)", ConsoleColor.Cyan);
        string spreadsheetId = Console.ReadLine();
        
        Console.WriteLine("Loading metadata from the spreadsheet...");
        Spreadsheet spreadsheet = sheetsManager.GetSpreadSheet(spreadsheetId);
        
        List<string> requestRanges = new List<string>();
        foreach (Sheet sheet in spreadsheet.Sheets)
        {
            requestRanges.Add($"{sheet.Properties.Title}!A:ZZZ");
        }
        
        Console.WriteLine("Loading cell data from the spreadsheet...");
        BatchGetValuesResponse valueRanges = sheetsManager.GetMultipleValues(spreadsheetId, requestRanges.ToArray());
        var spreadsheetValues = valueRanges.ValueRanges.ToList();
        var sheets = ValueRangesToStringLists(spreadsheetValues);

        ParsingOptions options = ObtainOptions(spreadsheet, sheets);

        MSBP msbp = ObtainMsbp(spreadsheet, sheets);

        List<string> langNames = AskLanguageNames(spreadsheet, sheets);

        List<List<MSBT>> langs = ObtainMsbts(spreadsheet, sheets, options, msbp, langNames);

        AskOutputPath(options);
        
        SaveMsbts(langs, msbp, options);
        
        ConsoleUtils.Exit();
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

    static ParsingOptions ObtainOptions(Spreadsheet spreadsheet, List<List<List<string>>> sheets)
    {
        int sheetId = spreadsheet.Sheets.ToList().FindIndex(x => x.Properties.Title == "#Settings");
        ParsingOptions options = new();
        
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
        }

        return options;
    }

    static MSBP ObtainMsbp(Spreadsheet spreadsheet, List<List<List<string>>> sheets)
    {
        MSBP msbp = new();

        ObtainMsbpColors(spreadsheet, sheets, msbp);
        ObtainMsbpStyles(spreadsheet, sheets, msbp);
        ObtainMsbpAttributes(spreadsheet, sheets, msbp);
        ObtainMsbpTags(spreadsheet, sheets, msbp);

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
                    msbp.TagGroups.Last().Tags.Last().Parameters.Last().List.Add(row[3]);
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

    static List<string> AskLanguageNames(Spreadsheet spreadsheet, List<List<List<string>>> sheets)
    {
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
            }

            break;
        }

        List<int> wantedLanguageIds = new();
        foreach (var name in wantedLanguageNames)
        {
            wantedLanguageIds.Add(headerRow.IndexOf(name));
        }

        return wantedLanguageNames;
    }

    static List<List<MSBT>> ObtainMsbts(Spreadsheet spreadsheet, List<List<List<string>>> sheets,
        ParsingOptions options, MSBP msbp, List<string> langNames)
    {
        Console.Clear();
        
        List<List<MSBT>> langs = new();
        for (int i = 0; i < langNames.Count; i++)
        {
            langs.Add(new());
        }
        
        var msbtSheetIds = GetMsbtSheetIds(spreadsheet, sheets);
        foreach (var msbtSheetId in msbtSheetIds)
        {
            var sheet = spreadsheet.Sheets[msbtSheetId];
            var sheetName = sheet.Properties.Title;
            var sheetGrid = sheets[msbtSheetId];
            var headerRow = sheetGrid[0];

            var internalSheetId = spreadsheet.Sheets.ToList().FindIndex(x => x.Properties.Title == "#InternalData");
            if (internalSheetId == -1)
            {
                ConsoleUtils.Exit("Your spreadsheet doesn't contain an #InternalData sheet.");
            }
            
            var internalDataRowId = sheets[internalSheetId].FindIndex(x => x[0] == sheetName);
            var internalDataRow = sheets[internalSheetId][internalDataRowId];
            var slotCount = Convert.ToUInt32(internalDataRow[1]);
            var version = Convert.ToByte(internalDataRow[2]);
            var byteOrder = internalDataRow[3] == "Little Endian" ? Endianness.LittleEndian : Endianness.BigEndian;
            Enum.TryParse(internalDataRow[4], out EncodingType encoding);
            var ato1 = new List<int>();
            var hasAto1 = spreadsheet.Sheets[internalSheetId].Properties.GridProperties.ColumnCount == 6;
            if (hasAto1)
            {
                ato1 = CommaSpaceNumbersToList(internalDataRow[5]);
            }
            
            foreach (var langName in langNames)
            {
                Console.WriteLine($"Obtaining {sheetName}.msbt ({langName})...");
                var langColumnId = headerRow.IndexOf(langName);
                
                var hasAtr1 = false;
                var hasTsy1 = false;
                var usesAttributeString = false;
                var attributeColumnId = -1;
                
                if (headerRow.Contains("Attributes"))
                {
                    attributeColumnId = headerRow.IndexOf("Attributes");
                }

                if (headerRow.Count(x => x.EndsWith('%')) > 0)
                {
                    attributeColumnId = langColumnId + 1;
                }

                if (sheet.Properties.GridProperties.RowCount == 1 && attributeColumnId != -1)
                {
                    hasAtr1 = true;
                    hasTsy1 = true;
                }

                bool noFileFlag = false;
                var messages = new Dictionary<object, Message>();
                for (int i = 1; i < sheetGrid.Count; i++)
                {
                    var messageRow = sheetGrid[i];

                    var label = messageRow[0];
                    var text = messageRow[langColumnId];
                    
                    if (text == "{{no-translation}}")
                    {
                        text = messageRow[1];
                    }
                    if (text == "{{no-message}}")
                    {
                        continue;
                    }
                    if (text == "{{no-file}}")
                    {
                        noFileFlag = true;
                        break;
                    }
                    
                    var attributeByteData = new byte[]{};
                    var attributeString = "";
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

                        if (attributeDict.Count > 0)
                        {
                            hasAtr1 = true;

                            if (attributeDict.ContainsKey("StringAttribute"))
                            {
                                usesAttributeString = true;
                                attributeString = attributeDict["StringAttribute"];
                                attributeDict.Remove("StringAttribute");
                            }
                            
                            if (attributeDict.Count == 1 && attributeDict.ContainsKey("Attributes"))
                            {
                                attributeByteData = Convert.FromHexString(attributeDict["Attributes"].Replace("-", ""));
                                attributeDict.Remove("Attributes");
                            }

                            attributeByteData = MessageAttribute.KeysAndValuesToBytes(attributeDict.Keys.ToList(),
                                attributeDict.Values.ToList(), msbp);
                        }
                    }
                    
                    messages.Add(label, new Message
                    {
                        Text = text,
                        StyleId = styleId,
                        Attribute = new MessageAttribute
                        {
                            ByteData = attributeByteData,
                            StringData = attributeString
                        }
                    });
                }

                uint bytesPerAttribute = 0;
                if (messages.Count != 0)
                {
                    bytesPerAttribute = (uint)messages.First().Value.Attribute.ByteData.Length;
                }
                foreach (var message in messages)
                {
                    int curLength = message.Value.Attribute.ByteData.Length;
                    if (curLength != bytesPerAttribute)
                    {
                        throw new InvalidDataException(
                            $"Atrribute lengths of messages \"{messages.First().Key}\" and \"{message.Key}\" mismatch ({bytesPerAttribute} and {curLength})");
                    }
                }
                
                langs[langNames.IndexOf(langName)].Add(new MSBT
                {
                    FileName = sheetName,
                    Language = langName,
                    Header = new Header()
                    {
                        FileType = FileType.MSBT,
                        Version = version,
                        Endianness = byteOrder,
                        EncodingType = encoding
                    },
                    LabelSlotCount = slotCount,
                    BytesPerAttribute = bytesPerAttribute,
                    UsesAttributeStrings = usesAttributeString,
                    ATO1Numbers = ato1,
                    HasLBL1 = true,
                    HasATR1 = hasAtr1,
                    HasTSY1 = hasTsy1,
                    HasATO1 = hasAto1,
                    Messages = messages
                });
            }
        }

        return langs;
    }
    
    static List<int> GetMsbtSheetIds(Spreadsheet spreadsheet, List<List<List<string>>> sheets)
    {
        List<int> ids = new();
        int counter = 0;
        foreach (var sheet in spreadsheet.Sheets)
        {
            if (!sheet.Properties.Title.StartsWith('#'))
            {
                ids.Add(counter);
            }

            counter++;
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

                value = cell[..FindFirstIndexWhereNotAfter(cell, '\\', '"')];
                cell = cell[value.Length..];
                value = value[1..(value.Length - 1)];
                value = value.Replace("\\\"", "\"");
                value = value.Replace("\\\\", "\\");
                
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

    static int FindFirstIndexWhereNotAfter(string str, char first, char second)
    {
        for (int i = 0; i < str.Length; i++)
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
        Console.Clear();
        Console.WriteLine("Enter the path for outputting MSBT files:");
        var path = Console.ReadLine();

        options.OutputPath = path;
    }

    static void SaveMsbts(List<List<MSBT>> langs, MSBP msbp, ParsingOptions options)
    {
        Console.Clear();

        foreach (var lang in langs)
        {
            foreach (var msbt in lang)
            {
                Console.WriteLine($"Saving {msbt.Language}\\{options.UnnecessaryPathPrefix}{msbt.FileName}.msbt");
                string filePath = $"{options.OutputPath}\\{msbt.Language}\\{options.UnnecessaryPathPrefix}{msbt.FileName}.msbt";
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                File.WriteAllBytes(filePath, msbt.Compile(options, msbp));
            }
        }
        
        Console.WriteLine("\nDone!");
        ConsoleUtils.Exit();
    }
}
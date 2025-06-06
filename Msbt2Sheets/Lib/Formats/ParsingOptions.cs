using System.Text;
using Msbt2Sheets.Lib.Formats.FileComponents;

namespace Msbt2Sheets.Lib.Formats;

public class ParsingOptions
{
    public bool ShortenTags = false;

    public bool ShortenPagebreak = false;

    public bool AddLinebreaksAfterPagebreaks = false;

    public bool SkipRuby = false;

    public int FreezeColumnCount = 2;

    public int ColumnSize = 250;

    public bool HighlightTags = false;

    public string UnnecessaryPathPrefix = "";

    public string ColorIdentification = "By RGBA bytes";
    
    public List<string> TransLangNames = new();

    public List<string> TransLangSheetNames = new();
    
    public List<string> TransLangPaths = new();

    //output

    public string SpreadsheetId = "";
    public string OutputPath = "";

    public bool NoStatsSheet = false;
    public bool NoSettingsSheet = false;
    public bool NoInternalDataSheet = false;

    public bool SkipLangIfNotTranslated = false;
    public bool ExtendedHeader = false;
    public List<string> SheetNames = new();
    public List<string> CustomFileNames = new();

    public byte GlobalVersion { get; set; }
    public Endianness GlobalEndianness { get; set; }
    public EncodingType GlobalEncodingType { get; set; }
    public List<int> GlobalAto1 = new();
    public List<uint> SlotCounts = new();

    public string NoTranslationSymbol = "{{no-translation}}";
    public string NoMessageSymbol = "{{no-message}}";
    public string NoFileSymbol = "{{no-file}}";
    public int MainLangColumnId = 1;

    public List<string> UiLangNames = new();
    public List<string> InternalLangNames = new();
    public List<string> OutputLangNames = new();

    public bool RecreateSources = false;
    
    public bool ChangeItemTagToValueTag = false;

    public bool UploadInParts = false;

    public int MessagesPerRequest = 50000;

    public ParsingOptions() { }

    public ParsingOptions(Dictionary<string, string> fileOptions)
    {
        foreach (var option in fileOptions)
        {
            switch (option.Key)
            {
                case "spreadsheetId":
                    SpreadsheetId = option.Value;
                    break;
                case "outputPath":
                    OutputPath = option.Value;
                    break;
                case "noStatsSheet":
                    NoStatsSheet = Convert.ToBoolean(option.Value);
                    break;
                case "noSettingsSheet":
                    NoSettingsSheet = Convert.ToBoolean(option.Value);
                    break;
                case "noInternalDataSheet":
                    NoInternalDataSheet = Convert.ToBoolean(option.Value);
                    break;
                case "addLinebreaksAfterPagebreaks":
                    AddLinebreaksAfterPagebreaks = Convert.ToBoolean(option.Value);
                    break;
                case "colorIdentification":
                    ColorIdentification = option.Value;
                    break;
                case "skipLangIfNotTranslated":
                    SkipLangIfNotTranslated = Convert.ToBoolean(option.Value);
                    break;
                case "extendedHeader":
                    ExtendedHeader = Convert.ToBoolean(option.Value);
                    break;
                case "sheetNames":
                    SheetNames = option.Value.Split('|').ToList();
                    break;
                case "customFileNames":
                    CustomFileNames = option.Value.Split('|').ToList();
                    break;
                case "globalVersion":
                    GlobalVersion = Convert.ToByte(option.Value);
                    break;
                case "globalEndianness":
                    GlobalEndianness = option.Value == "Little Endian" ? Endianness.LittleEndian : Endianness.BigEndian;
                    break;
                case "globalEncoding":
                    Enum.TryParse(option.Value, out EncodingType encoding);
                    GlobalEncodingType = encoding;
                    break;
                case "globalAto1":
                    string[] ato1sStr = option.Value.Split(", ");
                    foreach (var ato1Str in ato1sStr)
                    {
                        GlobalAto1.Add(Convert.ToInt32(ato1Str));
                    }
                    break;
                case "slotCounts":
                    string[] slotCountsStr = option.Value.Split('|');
                    foreach (var slotCountStr in slotCountsStr)
                    {
                        SlotCounts.Add(Convert.ToUInt32(slotCountStr));
                    }
                    break;
                case "uiLangs":
                    UiLangNames = option.Value.Split('|').ToList();
                    break;
                case "outputLangs":
                    OutputLangNames = option.Value.Split('|').ToList();
                    break;
                case "noTranslationSymbol":
                    NoTranslationSymbol = option.Value;
                    break;
                case "noMessageSymbol":
                    NoMessageSymbol = option.Value;
                    break;
                case "noFileSymbol":
                    NoFileSymbol = option.Value;
                    break;
                case "mainLangColumnId":
                    MainLangColumnId = Convert.ToInt32(option.Value);
                    break;
                case "recreateSources":
                    RecreateSources = Convert.ToBoolean(option.Value);
                    break;
                case "changeItemTagToValueTag":
                    ChangeItemTagToValueTag = Convert.ToBoolean(option.Value);
                    break;
                case "uploadInParts":
                    UploadInParts = Convert.ToBoolean(option.Value);
                    break;
                case "messagesPerRequest":
                    MessagesPerRequest = Convert.ToInt32(option.Value);
                    break;
            }
        }

        if (UiLangNames.Count > 0 && OutputLangNames.Count == 0)
        {
            OutputLangNames = UiLangNames;
        }
    }
}
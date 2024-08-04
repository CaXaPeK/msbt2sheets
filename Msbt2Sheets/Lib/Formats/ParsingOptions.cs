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
    public List<string> OutputLangNames = new();

    public ParsingOptions()
    {
    }
}
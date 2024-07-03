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
    
    public List<string> TransLangNames = new();
    public List<string> TransLangPaths = new();

    public ParsingOptions()
    {
    }
}
namespace Msbt2Sheets.Lib.Formats.FileComponents;

public class Style
{
    public string Name { get; set; }
    
    public int RegionWidth { get; set; }
    
    public int LineCount { get; set; }
    
    public int FontId { get; set; }
    
    public int BaseColorId { get; set; }
    
    public Style(int regionWidth, int lineCount, int fontId, int baseColorId)
    {
        RegionWidth = regionWidth;
        LineCount = lineCount;
        FontId = fontId;
        BaseColorId = baseColorId;
    }

    public Style(string name, int regionWidth, int lineCount, int fontId, int baseColorId)
    {
        Name = name;
        RegionWidth = regionWidth;
        LineCount = lineCount;
        FontId = fontId;
        BaseColorId = baseColorId;
    }
}
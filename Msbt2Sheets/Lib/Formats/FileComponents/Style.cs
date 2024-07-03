namespace Msbt2Sheets.Lib.Formats.FileComponents;

public class Style
{
    public string Name { get; set; }
    
    public uint RegionWidth { get; set; }
    
    public uint LineCount { get; set; }
    
    public uint FontId { get; set; }
    
    public uint BaseColorId { get; set; }
    
    public Style(uint regionWidth, uint lineCount, uint fontId, uint baseColorId)
    {
        RegionWidth = regionWidth;
        LineCount = lineCount;
        FontId = fontId;
        BaseColorId = baseColorId;
    }
}
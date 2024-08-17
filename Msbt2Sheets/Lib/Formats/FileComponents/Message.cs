namespace Msbt2Sheets.Lib.Formats.FileComponents;

public class Message
{
    public string Text { get; set; }
    public int StyleId { get; set; }
    public Dictionary<string, object> Attributes { get; set; }
    
    //TODO: add attribute info here to avoid looking up attributes types
}
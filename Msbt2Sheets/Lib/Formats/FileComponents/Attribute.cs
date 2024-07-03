namespace Msbt2Sheets.Lib.Formats.FileComponents;

public class MessageAttribute
{
    public byte[] ByteData;
    public string StringData;

    public MessageAttribute() {}

    public MessageAttribute(byte[] byteData)
    {
        ByteData = byteData;
    }
    public MessageAttribute(byte[] byteData, string stringData)
    {
        ByteData = byteData;
        StringData = stringData;
    }
}

public class AttributeInfo
{
    public string Name { get; set; }
    
    public ParamType Type { get; set; }

    public List<string> List = new();
    
    public uint Offset { get; set; }

    public AttributeInfo(string name, byte type, uint offset, List<string> list)
    {
        Name = name;
        Type = (ParamType)type;
        Offset = offset;
        List = list;
    }
}
using System.Text;
using Msbt2Sheets.Lib.Utils;

namespace Msbt2Sheets.Lib.Formats.FileComponents;

public class MessageAttribute
{
    public byte[] ByteData;
    public string? StringData = null;

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

    public List<string> ToStringList(MSBP? msbp)
    {
        List<string> list = new();
        
        if (msbp == null)
        {
            list.Add($"Attributes: {BitConverter.ToString(ByteData)}");
            if (StringData != null)
            {
                list.Add($"StringAttribute: {GeneralUtils.QuoteString(StringData)}");
            }

            return list;
        }

        if (!msbp.HasATI2)
        {
            list.Add($"Attributes: {BitConverter.ToString(ByteData)}");
            if (StringData != null)
            {
                list.Add($"StringAttribute: {GeneralUtils.QuoteString(StringData)}");
            }

            return list;
        }

        try
        {
            using (var reader = new BinaryReader(new MemoryStream(ByteData)))
            {
                for (int i = 0; i < msbp.AttributeInfos.Count; i++)
                {
                    var attrInfo = msbp.AttributeInfos[i];
                    if (attrInfo.Offset >= reader.BaseStream.Length)
                    {
                        break;
                    }

                    reader.BaseStream.Seek(attrInfo.Offset, SeekOrigin.Begin);

                    string value = "";

                    switch (attrInfo.Type)
                    {
                        case ParamType.UInt8:
                            value = reader.ReadByte().ToString();
                            break;
                        case ParamType.UInt16:
                            value = reader.ReadUInt16().ToString();
                            break;
                        case ParamType.UInt32:
                            value = reader.ReadUInt32().ToString();
                            break;
                        case ParamType.Int8:
                            value = ((sbyte) reader.ReadByte()).ToString();
                            break;
                        case ParamType.Int16:
                            value = reader.ReadInt16().ToString();
                            break;
                        case ParamType.Int32:
                            value = reader.ReadInt32().ToString();
                            break;
                        case ParamType.Float:
                            value = reader.ReadSingle().ToString();
                            break;
                        case ParamType.Double:
                            value = reader.ReadDouble().ToString();
                            break;
                        case ParamType.String:
                            //ushort strLength = reader.ReadUInt16();
                            //value = Encoding.Unicode.GetString(reader.ReadBytes(strLength));
                            value = reader.ReadUInt32().ToString();
                            break;
                        case ParamType.List:
                            byte id = reader.ReadByte();
                            value = attrInfo.List[id];
                            break;
                    }

                    list.Add($"{attrInfo.Name}: {value}");
                }
            }
        }
        catch
        {
            list.Add($"Attributes: {BitConverter.ToString(ByteData)}");
        }
        
        if (StringData != null)
        {
            list.Add($"StringAttribute: {GeneralUtils.QuoteString(StringData)}");
        }

        return list;
    }
    
    public static byte[] KeysAndValuesToBytes(List<string> keys, List<string> values, MSBP msbp)
    {
        List<byte> bytes = new();
        
        for (int i = 0; i < keys.Count; i++)
        {
            var attr = msbp.AttributeInfos.FirstOrDefault(x => x.Name == keys[i]);
            if (attr == null)
            {
                throw new InvalidDataException($"Can't parse attributes: no attribute with name {keys[i]} on the #Attributes sheet");
            }
            
            switch (attr.Type)
            {
                case ParamType.UInt8:
                    bytes.Add(Convert.ToByte(values[i]));
                    break;
                case ParamType.UInt16:
                    bytes.AddRange(BitConverter.GetBytes(Convert.ToUInt16(values[i])));
                    break;
                case ParamType.UInt32:
                    bytes.AddRange(BitConverter.GetBytes(Convert.ToUInt32(values[i])));
                    break;
                case ParamType.Int8:
                    bytes.Add((byte)Convert.ToSByte(values[i]));
                    break;
                case ParamType.Int16:
                    bytes.AddRange(BitConverter.GetBytes(Convert.ToInt16(values[i])));
                    break;
                case ParamType.Int32:
                    bytes.AddRange(BitConverter.GetBytes(Convert.ToInt32(values[i])));
                    break;
                case ParamType.Float:
                    bytes.AddRange(BitConverter.GetBytes(Convert.ToSingle(values[i])));
                    break;
                case ParamType.Double:
                    bytes.AddRange(BitConverter.GetBytes(Convert.ToDouble(values[i])));
                    break;
                case ParamType.String:
                    bytes.AddRange(BitConverter.GetBytes(Convert.ToUInt32(values[i])));
                    break;
                case ParamType.List:
                    int itemId = attr.List.IndexOf(values[i]);
                    if (itemId == -1)
                    {
                        throw new InvalidDataException(
                            $"Can't parse attributes: attribute \"{attr.Name}\" doesn't have a list item \"{values[i]}\"");
                    }
                    bytes.Add((byte)itemId);
                    break;
                default:
                    throw new InvalidDataException(
                        $"Can't parse attributes: attribute {attr.Name} has an invalid type of {(byte)attr.Type}");
            }
        }

        return bytes.ToArray();
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
    
    public AttributeInfo(string name, ParamType type, uint offset, List<string> list)
    {
        Name = name;
        Type = type;
        Offset = offset;
        List = list;
    }

    public object Read(FileReader reader, long attrStartPosition, long startPosition, Encoding encoding)
    {
        switch (Type)
        {
            case ParamType.UInt8:
                return reader.ReadByteAt(attrStartPosition);
            case ParamType.UInt16:
                return reader.ReadUInt16At(attrStartPosition);
            case ParamType.UInt32:
                return reader.ReadUInt32At(attrStartPosition);
            case ParamType.Int8:
                return reader.ReadSByteAt(attrStartPosition);
            case ParamType.Int16:
                return reader.ReadInt16At(attrStartPosition);
            case ParamType.Int32:
                return reader.ReadInt32At(attrStartPosition);
            case ParamType.Float:
                return reader.ReadSingleAt(attrStartPosition);
            case ParamType.Double:
                return reader.ReadDoubleAt(attrStartPosition);
            case ParamType.String:
                int stringOffset = reader.ReadInt32At(attrStartPosition);
                return reader.PeekTerminatedString(startPosition + stringOffset, encoding);
            case ParamType.List:
                byte itemId = reader.ReadByteAt(attrStartPosition);
                return List[itemId];
            default:
                throw new Exception($"Attribute {Name} has an unknown type!");
        }
    }
}
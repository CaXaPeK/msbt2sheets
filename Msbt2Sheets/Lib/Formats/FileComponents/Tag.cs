using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using Msbt2Sheets.Lib.Utils;

namespace Msbt2Sheets.Lib.Formats.FileComponents;

public class Tag
{
    public ushort Group { get; set; }
    public ushort Type { get; set; }
    public byte[] RawParameters { get; set; }
    public bool IsTagEnd { get; set; }

    public Tag(ushort group, ushort type, byte[] rawParameters, bool isTagEnd)
    {
        Group = group;
        Type = type;
        RawParameters = rawParameters;
        IsTagEnd = isTagEnd;
    }

    public string Stringify(ParsingOptions options, MSBP? msbp = null)
    {
        if (msbp == null)
        {
            if (IsTagEnd)
            {
                return $"</{Group}.{Type}>";
            }
            else
            {
                if (RawParameters.Length == 0)
                {
                    return $"<{Group}.{Type}>";
                }
                else
                {
                    return $"<{Group}.{Type}:{BitConverter.ToString(RawParameters)}>";
                }
            }
        }
        else
        {
            try
            {
                TagGroup group = msbp.TagGroups[Group];
                TagType type = msbp.TagGroups[Group].Tags[Type];

                if (Group == 0 && Type == 3)
                {
                    if (IsTagEnd)
                    {
                        if (options.ShortenTags)
                        {
                            return "</Color>";
                        }
                        else
                        {
                            return "</System.Color>";
                        }
                    }
                    else
                    {
                        var colorChannels = RawParameterToList(type, msbp);
                        Color color = Color.FromArgb((byte)colorChannels[3], (byte)colorChannels[0], (byte)colorChannels[1], (byte)colorChannels[2]);
                        string colorName = msbp.Colors.FirstOrDefault(x => x.Value == color).Key;
                    
                        if (options.ShortenTags)
                        {
                            return $"<Color {colorName}>";
                        }
                        else
                        {
                            return $"<System.Color {colorName}>";
                        }
                    }
                }

                if (Group == 0 && Type == 0 && options.SkipRuby)
                {
                    return "";
                }

                if (Group == 0 && Type == 4)
                {
                    if (options.AddLinebreaksAfterPagebreaks)
                    {
                        if (options.ShortenPagebreak) return "<p>\n";
                    
                        if (options.ShortenTags)
                        {
                            return "<PageBreak>\n";
                        }
                        return "<System.PageBreak>\n";
                    }

                    if (options.ShortenPagebreak) return "<p>";
                }

                if (IsTagEnd)
                {
                    if (options.ShortenTags)
                    {
                        return $"</{type.Name}>";
                    }
                    return $"</{group.Name}.{type.Name}>";
                }
                
                if (RawParameters.Length == 0)
                {
                    if (options.ShortenTags)
                    {
                        return $"<{type.Name}>";
                    }
                    return $"<{group.Name}.{type.Name}>";
                }
                
                if (options.ShortenTags)
                {
                    return $"<{type.Name}{RawParametersToString(type, msbp, options.ShortenTags)}>";
                }
                return $"<{group.Name}.{type.Name}{RawParametersToString(type, msbp, options.ShortenTags)}>";
            }
            catch
            {
                return this.Stringify(options);
            }
        }
    }

    private List<Object> RawParameterToList(TagType tag, MSBP msbp)
    {
        List<Object> list = new();

        if (tag.Name == "Ruby")
        {
            tag.Parameters = new List<TagParameter>()
            {
                new TagParameter()
                {
                    Name = "num",
                    Type = ParamType.UInt16
                },
                new TagParameter()
                {
                    Name = "rt",
                    Type = ParamType.String
                }
            };
        }
        
        int position = 0;
        foreach (var param in tag.Parameters)
        {
            if (position == RawParameters.Length)
            {
                break;
            }

            var paramType = param.Type;
            if (tag.Name == "Font")
            {
                paramType = ParamType.UInt16;
            }
            Object paramValue;
            switch (paramType)
            {
                case ParamType.UInt8:
                    paramValue = RawParameters[position];
                    position += 1;
                    break;
                case ParamType.UInt16:
                    paramValue = BitConverter.ToUInt16(RawParameters[position..(position + 2)]);
                    position += 2;
                    break;
                case ParamType.UInt32:
                    paramValue = BitConverter.ToUInt32(RawParameters[position..(position + 4)]);
                    position += 4;
                    break;
                case ParamType.Int8:
                    paramValue = (sbyte) RawParameters[position];
                    position += 1;
                    break;
                case ParamType.Int16:
                    paramValue = BitConverter.ToInt16(RawParameters[position..(position + 2)]);
                    position += 2;
                    break;
                case ParamType.Int32:
                    paramValue = BitConverter.ToInt32(RawParameters[position..(position + 4)]);
                    position += 4;
                    break;
                case ParamType.Float:
                    paramValue = BitConverter.ToSingle(RawParameters[position..(position + 4)]);
                    position += 4;
                    break;
                case ParamType.Double:
                    paramValue = BitConverter.ToDouble(RawParameters[position..(position + 8)]);
                    position += 8;
                    break;
                case ParamType.String:
                    ushort strLength = BitConverter.ToUInt16(RawParameters[position..(position + 2)]);
                    position += 2;
                    paramValue = Encoding.Unicode.GetString(RawParameters[position..(position + strLength)]);
                    paramValue = $"\"{paramValue}\"";
                    position += strLength;
                    break;
                case ParamType.List:
                    byte listItemId = RawParameters[position];
                    position += 2;
                    paramValue = param.List[param.ListItemIds.IndexOf(listItemId)];
                    break;
                default:
                    throw new InvalidDataException($"Tag \"{tag.Name}\" from MSBP has an invalid parameter type!");
            }

            list.Add(paramValue);
        }

        return list;
    }

    private string RawParametersToString(TagType tag, MSBP msbp, bool shorten = false)
    {
        string result = "";

        var parameters = RawParameterToList(tag, msbp);

        for (int i = 0; i < parameters.Count; i++)
        {
            if (!shorten)
            {
                result += $" {tag.Parameters[i].Name}={parameters[i]}";
            }
            else
            {
                result += $" {parameters[i]}";
            }
        }

        return result;
    }

    public static void Write(FileWriter writer, string tag, MSBP? msbp = null)
    {
        string unformattedTagPattern = @"<\/?\d+\.\d+(?::[0-9A-F]{2}(-[0-9A-F]{2})*)?>";
        bool unformatted = Regex.IsMatch(tag, unformattedTagPattern);

        try
        {
            if (unformatted)
            {
                WriteUnformattedTag(writer, tag);
            }
            else
            {
                if (msbp == null)
                {
                    throw new InvalidDataException("MSBP wasn't provided");
                }
                WriteFormattedTag(writer, tag, msbp);
            }
        }
        catch (Exception e)
        {
            throw new InvalidDataException($"Couldn't parse tag {tag}: {e.Message}");
        }
    }

    private static void WriteUnformattedTag(FileWriter writer, string tag)
    {
        bool isTagEnd = tag[1] == '/';
        
        if (isTagEnd)
        {
            writer.WriteUInt16(0xF);
        }
        else
        {
            writer.WriteUInt16(0xE);
        }
        
        string groupString = isTagEnd ? tag[2..tag.IndexOf('.')] : tag[1..tag.IndexOf('.')];
        ushort group = Convert.ToUInt16(groupString);
        writer.WriteUInt16(group);

        bool hasParameters = tag.Contains(':');
        string typeString = hasParameters
            ? tag[(tag.IndexOf('.') + 1)..tag.IndexOf(':')]
            : tag[(tag.IndexOf('.') + 1)..tag.IndexOf('>')];
        ushort type = Convert.ToUInt16(typeString);
        writer.WriteUInt16(type);

        if (hasParameters)
        {
            string parametersString = tag[(tag.IndexOf(':') + 1)..tag.IndexOf('>')];
            byte[] parametersBytes = Convert.FromHexString(parametersString.Replace("-", ""));
            ushort parametersLength = (ushort)parametersBytes.Length;
            writer.WriteUInt16(parametersLength);
            writer.WriteBytes(parametersBytes);
        }
        else if (!isTagEnd)
        {
            writer.WriteUInt16(0);
        }
    }

    private static void WriteFormattedTag(FileWriter writer, string tag, MSBP msbp)
    {
        bool isTagEnd = tag[1] == '/';
        
        if (isTagEnd)
        {
            writer.WriteUInt16(0xF);
            tag = tag[2..];
        }
        else
        {
            writer.WriteUInt16(0xE);
            tag = tag[1..];
        }

        bool hasParameters = tag.Contains(' ');
        bool shortened;
        if (hasParameters)
        {
            shortened = !tag[..tag.IndexOf(' ')].Contains('.');
        }
        else
        {
            shortened = !tag[..tag.IndexOf('>')].Contains('.');
        }

        TagGroup group;
        TagType type;
        
        if (shortened)
        {
            string typeName;
            if (hasParameters)
            {
                typeName = tag[..tag.IndexOf(' ')];
            }
            else
            {
                typeName = tag[..tag.IndexOf('>')];
            }

            group = msbp.TagGroups.FirstOrDefault(x => x.Tags.Any(tag => tag.Name == typeName));
            if (group == null)
            {
                throw new InvalidDataException($"Can't find groupId or tagId of {tag}!");
            }

            type = group.Tags.FirstOrDefault(x => x.Name == typeName);
        }
        else
        {
            string groupName = tag[..tag.IndexOf('.')];
            group = msbp.TagGroups.FirstOrDefault(x => x.Name == groupName);
            tag = tag[(tag.IndexOf('.') + 1)..];
            
            string typeName;
            if (hasParameters)
            {
                typeName = tag[..tag.IndexOf(' ')];
            }
            else
            {
                typeName = tag[..tag.IndexOf('>')];
            }

            type = group.Tags.FirstOrDefault(x => x.Name == typeName);
        }
        
        ushort groupId = (ushort)msbp.TagGroups.IndexOf(group);
        ushort typeId = (ushort)group.Tags.IndexOf(type);
        writer.WriteUInt16(groupId);
        writer.WriteUInt16(typeId);

        if (hasParameters)
        {
            string parametersString = tag[tag.IndexOf(' ')..];
            List<string> parameterValues = new();
            if (group.Name == "System" && type.Name == "Color" && parametersString.Count(x => x == ' ') != 4)
            {
                string colorString = parametersString[1..parametersString.IndexOf('>')];
                Color color = msbp.Colors[colorString];
                writer.WriteUInt16(4);
                writer.WriteByte(color.R);
                writer.WriteByte(color.G);
                writer.WriteByte(color.B);
                writer.WriteByte(color.A);
            }
            else
            {
                parameterValues = ParseParametersString(parametersString, shortened);
                byte[] parameterBytes = ParameterValuesToByteArray(parameterValues, type.Parameters, msbp);
                writer.WriteUInt16((ushort) parameterBytes.Length);
                writer.WriteBytes(parameterBytes);
            }
        }
        else if (!isTagEnd)
        {
            writer.WriteUInt16(0);
        }
    }

    private static List<string> ParseParametersString(string parametersString, bool shortened = false)
    {
        List<string> parameterValues = new();
        
        while (parametersString.Length > 1)
        {
            string valueString = "";
            bool insideString = false;
            bool reachedParamValue = shortened;
            
            int i;
            for (i = 1; !((parametersString[i] == ' ' || parametersString[i] == '>') && !insideString); i++)
            {
                if (!reachedParamValue)
                {
                    if (parametersString[i] == '=')
                    {
                        reachedParamValue = true;
                    }
                    continue;
                }

                if (parametersString[i] == '"')
                {
                    insideString = !insideString;
                    continue;
                }

                if (parametersString[i..(i+1)] == "\\\"" && insideString)
                {
                    valueString += '"';
                    i++;
                    continue;
                }

                valueString += parametersString[i];
            }

            parametersString = parametersString[i..];
            parameterValues.Add(valueString);
        }

        return parameterValues;
    }

    private static byte[] ParameterValuesToByteArray(List<string> values, List<TagParameter> parameters, MSBP msbp)
    {
        List<byte> bytes = new();

        for(int i = 0; i < values.Count; i++)
        {
            switch (parameters[i].Type)
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
                    bytes.AddRange(BitConverter.GetBytes((ushort)(values[i].Length * 2)));
                    //bytes.AddRange(msbp.Header.Encoding.GetBytes(values[i]));
                    bytes.AddRange(Encoding.Unicode.GetBytes(values[i]));
                    break;
                case ParamType.List:
                    int localListItemId = parameters[i].List.IndexOf(values[i]);
                    int globalListItemId = parameters[i].ListItemIds.IndexOf((ushort)localListItemId);
                    bytes.Add((byte)globalListItemId);
                    bytes.Add(0xCD);
                    break;
                default:
                    throw new InvalidDataException(
                        $"Parameter {parameters[i].Name} has an invalid type of {(byte)parameters[i].Type}!");
            }
        }
        
        return bytes.ToArray();
    }
}

public class TagGroup
{
    public string Name { get; set; }
    public List<TagType> Tags = new();
}

public class TagType
{
    public string Name { get; set; }
    public List<TagParameter> Parameters = new();
}

public class TagParameter
{
    public string Name { get; set; }
    public ParamType Type { get; set; }
    public List<string> List = new();
    public List<ushort> ListItemIds = new();
}
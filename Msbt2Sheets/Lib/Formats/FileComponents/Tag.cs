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

    public string Stringify(ParsingOptions options, string tagOrigin, Encoding encoding, MSBP? msbp = null, bool isBaseMsbp = false, bool parseCd = false)
    {
        if (msbp == null)
        {
            if (Group == 0)
            {
                return Stringify(options, tagOrigin, encoding, MSBP.BaseMSBP, true);
            }
            
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
                TagGroup? group = msbp.TagGroups.FirstOrDefault(x => x.Id == Group);
                TagType type = msbp.TagGroups[Group].Tags[Type];

                bool shortenTags = options.ShortenTags;
                if (shortenTags && !isTagNameFirst(type.Name, msbp))
                {
                    shortenTags = false;
                }
                if (options.ShortenPagebreak)
                {
                    msbp.TagGroups[0].Tags[4].Name = "p";
                }

                if (Group == 0 && Type == 3)
                {
                    if (IsTagEnd)
                    {
                        if (shortenTags)
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
                        if (RawParameters.Length == 2)
                        {
                            options.ColorIdentification = "By color ID";
                            msbp.TagGroups[0].Tags[3].Parameters[0].Type = ParamType.Int16;
                            short colorId = (short)RawParameterToList(type, encoding, msbp)[0];
                            msbp.TagGroups[0].Tags[3].Parameters[0].Type = ParamType.UInt8;

                            if (msbp.Colors.Count > colorId)
                            {
                                string colorName = GeneralUtils.GetColorNameFromId(colorId, msbp);
                                    
                                if (shortenTags)
                                {
                                    return $"<Color {colorName}>";
                                }
                                else
                                {
                                    return $"<System.Color {colorName}>";
                                }
                            }
                            else
                            {
                                if (shortenTags)
                                {
                                    return $"<Color {colorId}>";
                                }
                                else
                                {
                                    return $"<System.Color {colorId}>";
                                }
                            }
                        }
                        
                        var colorChannels = RawParameterToList(type, encoding, msbp);
                        Color color = Color.FromArgb((byte)colorChannels[3], (byte)colorChannels[0], (byte)colorChannels[1], (byte)colorChannels[2]);

                        if (msbp.Colors.Count(x => x.Value == color) > 0 && msbp.HasCLB1)
                        {
                            var colorName = msbp.Colors.FirstOrDefault(x => x.Value == color).Key;
                        
                            if (shortenTags)
                            {
                                return $"<Color {colorName}>";
                            }
                            else
                            {
                                return $"<System.Color {colorName}>";
                            }
                        }
                        else
                        {
                            var colorString = GeneralUtils.ColorToString(color);
                            
                            if (shortenTags)
                            {
                                return $"<Color {colorString}>";
                            }
                            else
                            {
                                return $"<System.Color {colorString}>";
                            }
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
                    
                        if (shortenTags)
                        {
                            return "<PageBreak>\n";
                        }
                        return "<System.PageBreak>\n";
                    }

                    if (options.ShortenPagebreak) return "<p>";
                }

                if (IsTagEnd)
                {
                    if (shortenTags)
                    {
                        return $"</{type.Name}>";
                    }
                    return $"</{group.Name}.{type.Name}>";
                }
                
                if (RawParameters.Length == 0)
                {
                    if (shortenTags)
                    {
                        return $"<{type.Name}>";
                    }
                    return $"<{group.Name}.{type.Name}>";
                }
                
                if (shortenTags)
                {
                    return $"<{type.Name}{RawParametersToString(type, msbp, encoding, options.ShortenTags, parseCd)}>";
                }
                return $"<{group.Name}.{type.Name}{RawParametersToString(type, msbp, encoding, options.ShortenTags, parseCd)}>";
            }
            catch
            {
                if (parseCd == false)
                {
                    try
                    {
                        return Stringify(options, tagOrigin, encoding, msbp, false, true);
                    }
                    catch
                    {
                        if (!isBaseMsbp)
                        {
                            Console.WriteLine($"Warning: Couldn't humanify the tag {Stringify(options, tagOrigin, encoding)} on {tagOrigin}.");
                        }

                        try
                        {
                            return Stringify(options, tagOrigin, encoding);
                        }
                        catch
                        {
                            throw new InvalidDataException($"Couldn't parse a tag on {tagOrigin} at all!");
                        }
                        
                    }
                }
                else
                {
                    throw new InvalidDataException();
                }
            }
        }
    }

    private bool isTagNameFirst(string tagName, MSBP msbp)
    {
        int occurrences = 0;
        foreach (var group in msbp.TagGroups)
        {
            foreach (var tag in group.Tags)
            {
                if (tag.Name == tagName)
                {
                    occurrences++;
                    
                    if (occurrences > 1)
                    {
                        return false;
                    }
                    
                    if (group.Id == Group)
                    {
                        return true;
                    }
                }
            }
        }

        return true;
    }

    private bool MsbpHasTag(string tagName, MSBP msbp)
    {
        foreach (var group in msbp.TagGroups)
        {
            foreach (var tag in group.Tags)
            {
                if (tag.Name == tagName)
                {
                    return true;
                }
            }
        }

        return false;
    }

    private List<Object> RawParameterToList(TagType tag, Encoding encoding, MSBP msbp, bool parseCd = false)
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
            if (position >= RawParameters.Length)
            {
                break;
            }

            if (RawParameters[position] == 0xCD && parseCd)
            {
                list.Add("CD");
                position++;
                if (position >= RawParameters.Length)
                {
                    break;
                }
            }

            var paramType = param.Type;
            if (Group == 0 && Type == 1)
            {
                paramType = ParamType.Int16;
            }
            Object paramValue;
            byte[] bytes;
            switch (paramType)
            {
                case ParamType.UInt8:
                    paramValue = RawParameters[position];
                    position += 1;
                    break;
                case ParamType.UInt16:
                    bytes = RawParameters[position..(position + 2)];
                    if (encoding == Encoding.BigEndianUnicode)
                    {
                        bytes = bytes.Reverse().ToArray();
                    }
                    paramValue = BitConverter.ToUInt16(bytes);
                    position += 2;
                    break;
                case ParamType.UInt32:
                    bytes = RawParameters[position..(position + 4)];
                    if (encoding == Encoding.BigEndianUnicode)
                    {
                        bytes = bytes.Reverse().ToArray();
                    }
                    paramValue = BitConverter.ToUInt32(bytes);
                    position += 4;
                    break;
                case ParamType.Int8:
                    paramValue = (sbyte) RawParameters[position];
                    position += 1;
                    break;
                case ParamType.Int16:
                    bytes = RawParameters[position..(position + 2)];
                    if (encoding == Encoding.BigEndianUnicode)
                    {
                        bytes = bytes.Reverse().ToArray();
                    }
                    paramValue = BitConverter.ToInt16(bytes);
                    position += 2;
                    break;
                case ParamType.Int32:
                    bytes = RawParameters[position..(position + 4)];
                    if (encoding == Encoding.BigEndianUnicode)
                    {
                        bytes = bytes.Reverse().ToArray();
                    }
                    paramValue = BitConverter.ToInt32(bytes);
                    position += 4;
                    break;
                case ParamType.Float:
                    bytes = RawParameters[position..(position + 4)];
                    if (encoding == Encoding.BigEndianUnicode)
                    {
                        bytes = bytes.Reverse().ToArray();
                    }
                    paramValue = BitConverter.ToSingle(bytes);
                    position += 4;
                    break;
                case ParamType.Double:
                    bytes = RawParameters[position..(position + 8)];
                    if (encoding == Encoding.BigEndianUnicode)
                    {
                        bytes = bytes.Reverse().ToArray();
                    }
                    paramValue = BitConverter.ToDouble(bytes);
                    position += 8;
                    break;
                case ParamType.String:
                    bytes = RawParameters[position..(position + 2)];
                    if (encoding == Encoding.BigEndianUnicode)
                    {
                        bytes = bytes.Reverse().ToArray();
                    }
                    ushort strLength = BitConverter.ToUInt16(bytes);
                    position += 2;
                    paramValue = encoding.GetString(RawParameters[position..(position + strLength)]);
                    paramValue = GeneralUtils.AddQuotesToString((string)paramValue);
                    position += strLength;
                    break;
                case ParamType.List:
                    byte listItemId = RawParameters[position];
                    position += 1;
                    //paramValue = param.List[param.ListItemIds.IndexOf(listItemId)];
                    paramValue = param.List[listItemId];
                    break;
                default:
                    throw new InvalidDataException($"Tag \"{tag.Name}\" from MSBP has an invalid parameter type!");
            }

            list.Add(paramValue);
        }

        if (position < RawParameters.Length)
        {
            list.Add(BitConverter.ToString(RawParameters[position..RawParameters.Length]));
        }

        return list;
    }

    private string RawParametersToString(TagType tag, MSBP msbp, Encoding encoding, bool shorten = false, bool parseCd = false)
    {
        string result = "";

        var parameters = RawParameterToList(tag, encoding, msbp, parseCd);

        for (int i = 0; i < parameters.Count; i++)
        {
            if (!shorten)
            {
                if (parameters[i].ToString() == "CD" && parseCd)
                {
                    result += $" bytes=CD";
                }
                else
                {
                    result += $" {tag.Parameters[i].Name}={parameters[i]}";
                }
            }
            else
            {
                result += $" {parameters[i]}";
            }
        }

        return result;
    }

    public static byte[] Write(FileWriter writer, string tag, ParsingOptions options, MSBP? msbp = null)
    {
        string unformattedTagPattern = @"<\/?\d+\.\d+(?::[0-9A-F]{2}(-[0-9A-F]{2})*)?>";
        bool unformatted = Regex.IsMatch(tag, unformattedTagPattern);

        try
        {
            if (unformatted)
            {
                return WriteUnformattedTag(writer, tag);
            }
            else
            {
                if (msbp == null)
                {
                    throw new InvalidDataException("MSBP wasn't provided");
                }
                return WriteFormattedTag(writer, tag, msbp, options);
            }
        }
        catch (Exception e)
        {
            throw new InvalidDataException($"Couldn't parse tag {tag}: {e.Message}");
        }
    }

    private static byte[] WriteUnformattedTag(FileWriter writer, string tag)
    {
        FileWriter tagWriter = new(new MemoryStream());
        bool isTagEnd = tag[1] == '/';
        
        if (isTagEnd)
        {
            writer.WriteUInt16(0xF);
            tagWriter.WriteUInt16(0xF);
        }
        else
        {
            writer.WriteUInt16(0xE);
            tagWriter.WriteUInt16(0xE);
        }
        
        string groupString = isTagEnd ? tag[2..tag.IndexOf('.')] : tag[1..tag.IndexOf('.')];
        ushort group = Convert.ToUInt16(groupString);
        writer.WriteUInt16(group);
        tagWriter.WriteUInt16(group);

        bool hasParameters = tag.Contains(':');
        string typeString = hasParameters
            ? tag[(tag.IndexOf('.') + 1)..tag.IndexOf(':')]
            : tag[(tag.IndexOf('.') + 1)..tag.IndexOf('>')];
        ushort type = Convert.ToUInt16(typeString);
        writer.WriteUInt16(type);
        tagWriter.WriteUInt16(type);

        if (hasParameters)
        {
            string parametersString = tag[(tag.IndexOf(':') + 1)..tag.IndexOf('>')];
            byte[] parametersBytes = Convert.FromHexString(parametersString.Replace("-", ""));
            ushort parametersLength = (ushort)parametersBytes.Length;
            writer.WriteUInt16(parametersLength);
            tagWriter.WriteUInt16(parametersLength);
            writer.WriteBytes(parametersBytes);
            tagWriter.WriteBytes(parametersBytes);
        }
        else if (!isTagEnd)
        {
            writer.WriteUInt16(0);
            tagWriter.WriteUInt16(0);
        }
        
        return ((MemoryStream)tagWriter.BaseStream).ToArray();
    }

    private static byte[] WriteFormattedTag(FileWriter writer, string tag, MSBP msbp, ParsingOptions options)
    {
        FileWriter tagWriter = new(new MemoryStream());
        string initialTag = tag;
        bool isTagEnd = tag[1] == '/';
        
        if (isTagEnd)
        {
            writer.WriteUInt16(0xF);
            tagWriter.WriteUInt16(0xF);
            tag = tag[2..];
        }
        else
        {
            writer.WriteUInt16(0xE);
            tagWriter.WriteUInt16(0xE);
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
            string typeName = hasParameters ? tag[..tag.IndexOf(' ')] : tag[..tag.IndexOf('>')];

            group = msbp.TagGroups.FirstOrDefault(x => x.Tags.Any(tag => tag.Name == typeName));
            if (group == null)
            {
                if (typeName == "p")
                {
                    group = msbp.TagGroups[0];
                    type = msbp.TagGroups[0].Tags[4];
                }
                else
                {
                    throw new InvalidDataException($"Can't find groupId of {initialTag}!");
                }
            }
            else
            {
                type = group.Tags.FirstOrDefault(x => x.Name == typeName);
                if (type == null)
                {
                    throw new InvalidDataException($"Can't find tagId of {initialTag} in {group.Name}!");
                }
            }
        }
        else
        {
            string groupName = tag[..tag.IndexOf('.')];
            group = msbp.TagGroups.FirstOrDefault(x => x.Name == groupName);
            tag = tag[(tag.IndexOf('.') + 1)..];
            
            string typeName = hasParameters ? tag[..tag.IndexOf(' ')] : tag[..tag.IndexOf('>')];
            
            type = group.Tags.FirstOrDefault(x => x.Name == typeName);
            if (type == null)
            {
                if (groupName == "System" && typeName == "p")
                {
                    type = msbp.TagGroups[0].Tags[4];
                }
                else
                {
                    throw new InvalidDataException($"Can't find tagId of {initialTag}!");
                }
            }
            else
            {
                type = group.Tags.FirstOrDefault(x => x.Name == typeName);
            }
        }
        
        ushort groupId = (ushort)msbp.TagGroups.IndexOf(group);
        ushort typeId = (ushort)group.Tags.IndexOf(type);
        writer.WriteUInt16(groupId);
        writer.WriteUInt16(typeId);
        tagWriter.WriteUInt16(groupId);
        tagWriter.WriteUInt16(typeId);

        if (hasParameters)
        {
            var parametersString = tag[tag.IndexOf(' ')..];
            List<string> parameterValues = new();
            if (group.Name == "System" && type.Name == "Color")
            {
                var colorString = parametersString[1..parametersString.IndexOf('>')];
                var color = new Color();
                short colorId = -1;
                if (msbp.Colors.ContainsKey(colorString))
                {
                    color = msbp.Colors[colorString];
                    colorId = (short)msbp.Colors.Keys.ToList().IndexOf(colorString);
                }
                else if (colorString.StartsWith('#'))
                {
                    colorString = colorString[1..];
                    var bytes = Convert.FromHexString(colorString);
                    color = Color.FromArgb(bytes[3], bytes[0], bytes[1], bytes[2]);
                }
                else
                {
                    colorId = Convert.ToInt16(colorString);
                }

                if (options.ColorIdentification == "By RGBA bytes")
                {
                    writer.WriteUInt16(4);
                    writer.WriteByte(color.R);
                    writer.WriteByte(color.G);
                    writer.WriteByte(color.B);
                    writer.WriteByte(color.A);
                    tagWriter.WriteUInt16(4);
                    tagWriter.WriteByte(color.R);
                    tagWriter.WriteByte(color.G);
                    tagWriter.WriteByte(color.B);
                    tagWriter.WriteByte(color.A);
                }
                else
                {
                    writer.WriteUInt16(2);
                    writer.WriteInt16(colorId);
                    tagWriter.WriteUInt16(2);
                    tagWriter.WriteInt16(colorId);
                }
            }
            else
            {
                parameterValues = ParseParametersString(parametersString, shortened);
                byte[] parameterBytes = ParameterValuesToByteArray(parameterValues, type.Parameters, msbp);
                writer.WriteUInt16((ushort) parameterBytes.Length);
                writer.WriteBytes(parameterBytes);
                tagWriter.WriteUInt16((ushort) parameterBytes.Length);
                tagWriter.WriteBytes(parameterBytes);
            }
        }
        else if (!isTagEnd)
        {
            writer.WriteUInt16(0);
            tagWriter.WriteUInt16(0);
        }
        
        return ((MemoryStream)tagWriter.BaseStream).ToArray();
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

        for (int i = 0; i < values.Count; i++)
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
    public ushort Id { get; set; }
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
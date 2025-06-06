using System.Data;
using System.Reflection.Emit;
using System.Text;
using Google.Apis.Sheets.v4.Data;
using Msbt2Sheets.Lib.Formats.FileComponents;
using Msbt2Sheets.Lib.Utils;

namespace Msbt2Sheets.Lib.Formats;

public class MSBT : GeneralFile
{
    public Dictionary<object, Message> Messages = new();
    public string FileName { get; set; }
    public string Language { get; set; }

    public bool HasLBL1 = false;
    public bool HasNLI1 = false;
    public bool HasATO1 = false;
    public bool HasATR1 = false;
    public bool HasTSY1 = false;

    //public bool UsesAttributeStrings { get; set; }

    //public uint BytesPerAttribute { get; set; }
    
    public uint LabelSlotCount { get; set; }

    public int MsbpAttributeCount = -1; //needed to recreate ATO1 if there's no MSBP

    public Header Header = new();
    
    public MSBT() {}
    
    public MSBT(Stream fileStream, ParsingOptions options, string? fileName = null, string? language = null, MSBP? msbp = null)
    {
        FileName = fileName.Substring(options.UnnecessaryPathPrefix.Length);
        Language = language;

        MemoryStream ms = new MemoryStream();
        fileStream.CopyTo(ms);
        FileReader reader = new(ms);
        
        LBL1 lbl1 = new();
        NLI1 nli1 = new();
        ATO1 ato1 = new();
        ATR1 atr1 = new();
        TSY1 tsy1 = new();
        TXT2 txt2 = new();
        
        Header = new(reader);

        for (int i = 0; i < Header.SectionCount; i++)
        {
            string sectionMagic = reader.ReadString(4, Encoding.ASCII);
            uint sectionSize = reader.ReadUInt32();
            reader.Skip(8);
            long startPosition = reader.Position;
            
            switch (sectionMagic)
            {
                case "LBL1":
                    HasLBL1 = true;
                    lbl1 = new(reader);
                    break;
                case "NLI1":
                    HasNLI1 = true;
                    nli1 = new(reader);
                    break;
                case "ATO1":
                    HasATO1 = true;
                    ato1 = new(reader, sectionSize);
                    break;
                case "ATR1":
                    HasATR1 = true;
                    atr1 = new(reader, ato1.Offsets, sectionSize, Header.Encoding, msbp != null ? msbp.AttributeInfos : new());
                    break;
                case "TSY1":
                    HasTSY1 = true;
                    tsy1 = new(reader, sectionSize);
                    break;
                case "TXT2":
                    txt2 = new(reader, HasATR1, atr1, HasTSY1, tsy1, options, lbl1, HasNLI1, nli1, fileName, Header.Encoding, msbp);
                    break;
                default:
                    throw new DataException($"Unknown section magic!");
            }
            
            reader.JumpTo(startPosition);
            reader.Skip((int)sectionSize);
            reader.Align(0x10);
        }

        if (HasLBL1)
        {
            LabelSlotCount = lbl1.LabelSlotCount;
            for (int i = 0; i < lbl1.Labels.Count; i++)
            {
                Messages.Add(lbl1.Labels[i], txt2.Messages[i]);
            }
        }
        else if (HasNLI1)
        {
            foreach (var nli1Entry in nli1.Indices)
            {
                Messages.Add(nli1Entry.Value.ToString(), txt2.Messages[(int)nli1Entry.Key]);
            }
        }

        if (HasATO1)
        {
            MsbpAttributeCount = ato1.Offsets.Count;
        }
        else if (msbp != null)
        {
            MsbpAttributeCount = msbp.AttributeInfos.Count;
        }
        else if (atr1.AttributeDicts.Count > 0)
        {
            var attrDict = atr1.AttributeDicts.First();
            if (!attrDict.ContainsKey("Attributes"))
            {
                if (attrDict.ContainsKey("StringAttributes"))
                {
                    MsbpAttributeCount = attrDict.Count - 1;
                }
                else
                {
                    MsbpAttributeCount = attrDict.Count;
                }
            }
        }
    }

    public byte[] Compile(ParsingOptions options, MSBP? msbp = null)
    {
        ushort sectionCount = 0;
        FileWriter writer = new(new MemoryStream());
        writer.Endianness = Header.Endianness;
        
        Header.Write(writer);
        
        if (HasLBL1)
        {
            LBL1.Write(writer, Messages, LabelSlotCount);
            sectionCount++;
        }
        
        if (HasATO1)
        {
            try
            {
                if (Messages.Count > 0)
                {
                    try
                    {
                        if (!Messages.First().Value.Attributes.ContainsKey("Attributes"))
                        {
                            ATO1.Write(writer, Messages.First().Value.Attributes, msbp != null ? msbp.AttributeInfos : new(), MsbpAttributeCount);
                            sectionCount++;
                        }
                    }
                    catch (Exception e)
                    {
                        throw new Exception(e.Message);
                    }
                }
                else
                {
                    throw new Exception("No messages."); //how can there be an ATO1 section when there's no messages?!?!?
                }
            }
            catch (Exception e)
            {
                throw new Exception($"Can't build an ATO1 section for {FileName}: {e.Message}.");
            }
        }
        
        if (HasATR1)
        {
            try
            {
                ATR1.Write(writer, Messages.Values.ToArray(), msbp != null ? msbp.AttributeInfos : new(), Header.Encoding);
            }
            catch (Exception e)
            {
                throw new Exception($"Can't build an ATR1 section for {FileName}: {e.Message}.");
            }
            
            sectionCount++;
        }
        
        if (HasTSY1)
        {
            TSY1.Write(writer, Messages.Values.ToArray());
            sectionCount++;
        }
        
        TXT2.Write(writer, Messages.Values.ToArray(), Header.Encoding, options, msbp);
        sectionCount++;
        
        CalculateAndSetSectionCountAndFileSize(writer, sectionCount, (uint)writer.Position);
        return ((MemoryStream)writer.BaseStream).ToArray();
    }

    #region msbt sections
    
    internal class LBL1
    {
        public List<String> Labels { get; set; }
        public uint LabelSlotCount { get; set; }

        public LBL1() {}

        public LBL1(FileReader reader)
        {
            long startPos = reader.Position;
            LabelSlotCount = reader.ReadUInt32();
            reader.JumpTo(startPos);
            
            Labels = ReadLabels(reader);
        }

        public static void Write(FileWriter writer, Dictionary<object, Message> messages, uint hashSlotCount)
        {
            writer.WriteString("LBL1", Encoding.ASCII);
            long sizePosition = writer.Position;
            writer.Pad(0xC);

            writer.WriteUInt32(hashSlotCount);

            Dictionary<uint, List<string>> hashTable = new();
            foreach (var messagePair in messages)
            {
                string label = messagePair.Key.ToString();
                uint hash = CalculateLabelHash(label, hashSlotCount);
                if (!hashTable.ContainsKey(hash))
                {
                    hashTable.Add(hash, new List<string>());
                }
                
                hashTable[hash].Add(label);
            }

            Dictionary<uint, List<string>> orderedHashTable = hashTable.OrderBy(x => x.Key)
                .ToDictionary(x => x.Key, x => x.Value);

            uint offsetToLables = hashSlotCount * 8 + 4;
            for (uint i = 0; i < hashSlotCount; i++)
            {
                if (orderedHashTable.ContainsKey(i))
                {
                    writer.WriteUInt32((uint)orderedHashTable[i].Count);
                    writer.WriteUInt32(offsetToLables);

                    foreach (string label in orderedHashTable[i])
                    {
                        offsetToLables += 1 + (uint)label.Length + 4;
                    }
                }
                else
                {
                    writer.WriteUInt32(0);
                    writer.WriteUInt32(offsetToLables);
                }
            }

            foreach (var entry in orderedHashTable)
            {
                foreach (var label in entry.Value)
                {
                    uint labelIndex = 0;
                    uint i = 0;
                    foreach (var message in messages)
                    {
                        if (message.Key.ToString() == label)
                        {
                            labelIndex = i;
                            break;
                        }

                        i++;
                    }
                    
                    writer.WriteByte((byte)label.Length);
                    writer.WriteString(label, Encoding.ASCII);
                    writer.WriteUInt32(labelIndex);
                }
            }

            CalculateAndSetSectionSize(writer, sizePosition);
            writer.Align(0x10, 0xAB);
        }
    }
    internal class NLI1
    {
        public Dictionary<uint, uint> Indices { get; set; }

        public NLI1() {}
        
        public NLI1(FileReader reader)
        {
            uint entryCount = reader.ReadUInt32();
            Indices = new();
            for (uint i = 0; i < entryCount; i++)
            {
                uint messageIndex = reader.ReadUInt32();
                uint entryId = reader.ReadUInt32();
                Indices.Add(entryId, messageIndex);
            }
        }
    }
    internal class ATO1
    {
        public List<int> Offsets = new();

        public ATO1() {}

        public ATO1(FileReader reader, long sectionSize)
        {
            Offsets = new();
            
            for (long i = 0; i < sectionSize / 4; i++)
            {
                Offsets.Add(reader.ReadInt32());
            }
        }

        public static void Write(FileWriter writer, Dictionary<string, object> attrDict, List<AttributeInfo> attrInfos, int msbpAttrCount)
        {
            if (attrDict.ContainsKey("Attributes"))
            {
                //throw new Exception("No distinct attributes.");
                return;
            }
            
            writer.WriteString("ATO1", Encoding.ASCII);
            long sizePosition = writer.Position;
            writer.Pad(0xC);

            int unnamedAttrCount = attrDict.Count(x => x.Key.StartsWith("Attribute_"));
            List<int> attrOffsets = Enumerable.Repeat(-1, msbpAttrCount).ToList();
            int curOffset = 0;
            
            if (!attrDict.ContainsKey("Attributes") && unnamedAttrCount == 0)
            {
                foreach (var attr in attrDict)
                {
                    int attrId = attrInfos.FindIndex(x => x.Name == attr.Key);
                    if (attrId == -1)
                    {
                        throw new Exception($"Can't find attribute \"{attr.Key}\".");
                    }

                    AttributeInfo attrInfo = attrInfos[attrId];
                    attrOffsets[attrId] = curOffset;
                    curOffset += GeneralUtils.GetTypeSize(attrInfo.Type);
                }
            }

            if (unnamedAttrCount > 0)
            {
                foreach (var attrEntry in attrDict)
                {
                    if (attrEntry.Key.StartsWith("Attribute_"))
                    {
                        int attrId = Convert.ToInt32(attrEntry.Key[(attrEntry.Key.IndexOf('_') + 1)..]);
                        if (attrEntry.Value is byte[] byteAttr)
                        {
                            attrOffsets[attrId] = curOffset;
                            curOffset += byteAttr.Length;
                        }
                    }
                }
            }
            
            foreach (var attrOffset in attrOffsets)
            {
                writer.WriteInt32(attrOffset);
            }
            
            CalculateAndSetSectionSize(writer, sizePosition);
            writer.Align(0x10, 0xAB);
        }
    }
    internal class ATR1
    {
        public List<Dictionary<string, object>> AttributeDicts = new();

        public ATR1() {}

        public ATR1(FileReader reader, List<int> attributeOffsets, uint sectionSize, Encoding encoding, List<AttributeInfo> attributeInfos)
        {
            long startPosition = reader.Position;
            long attrDataStartPosition = startPosition + 8;
            uint attributeCount = reader.ReadUInt32();
            int bytesPerAttribute = reader.ReadInt32();

            bool hasStringAttributes = sectionSize - 8 > attributeCount * bytesPerAttribute;
            long stringAttributesStartPosition = startPosition + 8 + attributeCount * bytesPerAttribute;
            List<string> stringAttributes = new();

            for (uint i = 0; i < attributeCount; i++)
            {
                Dictionary<string, object> attrDict = new();
                long attrsStartPosition = attrDataStartPosition + i * bytesPerAttribute;
                
                if (attributeInfos.Count > 0)
                {
                    if (attributeOffsets.Count > 0)
                    {
                        for (int j = 0; j < attributeOffsets.Count; j++)
                        {
                            int offset = attributeOffsets[j];
                            if (offset == -1)
                            {
                                continue;
                            }

                            long attrStartPosition = attrsStartPosition + offset;
                            AttributeInfo attrInfo = attributeInfos[j];

                            attrDict.Add(attrInfo.Name, attrInfo.Read(reader, attrStartPosition, startPosition, encoding));
                        }
                    }
                    else
                    {
                        int attrNum = 0;
                        while (reader.Position - attrsStartPosition < bytesPerAttribute)
                        {
                            AttributeInfo attrInfo = attributeInfos[attrNum];
                            attrDict.Add(attrInfo.Name, attrInfo.Read(reader, reader.Position, startPosition, encoding));
                            attrNum++;
                        }
                    }
                }

                if (attributeInfos.Count == 0)
                {
                    if (attributeOffsets.Count > 0)
                    {
                        List<int> attrIds = new();
                        for (int j = 0; j < attributeOffsets.Count; j++)
                        {
                            int offset = attributeOffsets[j];
                            if (offset != -1)
                            {
                                attrIds.Add(j);
                            }
                        }

                        for (int j = 0; j < attrIds.Count; j++)
                        {
                            int attrSize = j < attrIds.Count - 1
                                ? attributeOffsets[attrIds[j + 1]] - attributeOffsets[attrIds[j]]
                                : bytesPerAttribute - attributeOffsets[attrIds[j]];

                            byte[] attrBytes = reader.ReadBytes(attrSize);
                            string attrBytesStringified = BitConverter.ToString(attrBytes);
                        
                            attrDict.Add($"Attribute_{attrIds[j]}", attrBytesStringified);
                        }
                    }
                    else
                    {
                        byte[] attrBytes = reader.ReadBytes(bytesPerAttribute);
                        attrDict.Add("Attributes", attrBytes);
                    }

                    if (hasStringAttributes)
                    {
                        if (stringAttributes.Count == 0)
                        {
                            reader.JumpTo(stringAttributesStartPosition);
                            while (reader.Position < startPosition + sectionSize)
                            {
                                stringAttributes.Add(reader.ReadTerminatedString(encoding));
                            }
                        }

                        int stringAttributesPerMessage = stringAttributes.Count / (int)attributeCount;
                        List<string> curStringAttributes = new();
                        for (int j = 0; j < stringAttributesPerMessage; j++)
                        {
                            curStringAttributes.Add(stringAttributes[(int)i * stringAttributesPerMessage + j]);
                        }

                        /*string curStringAttributesStringified = "";
                        foreach (var str in curStringAttributes)
                        {
                            string quotedStr = GeneralUtils.QuoteString(str);
                            string addedStr = curStringAttributesStringified == "" ? quotedStr : $", {quotedStr}";
                            curStringAttributesStringified += addedStr;
                        }
                        curStringAttributesStringified += ';';*/
                        
                        attrDict.Add("StringAttributes", curStringAttributes);
                    }
                }
                
                AttributeDicts.Add(attrDict);
            }
        }

        public static void Write(FileWriter writer, Message[] messages, List<AttributeInfo> attrInfos, Encoding encoding)
        {
            writer.WriteString("ATR1", Encoding.ASCII);
            long sizePosition = writer.Position;
            writer.Pad(0xC);

            long startPosition = writer.Position;
            writer.WriteInt32(messages.Length);
            long bytesPerAttrPosition = writer.Position;
            writer.Pad(4);

            int bytesPerAttr = 0;
            List<string> stringAttrs = new();
            List<long> stringAttrOffsetPositions = new();
            for (int i = 0; i < messages.Length; i++)
            {
                int curBytesPerAttr = 0;
                foreach (var attrEntry in messages[i].Attributes)
                {
                    if (attrEntry.Value is byte[] attrBytes)
                    {
                        writer.WriteBytes(attrBytes);
                        curBytesPerAttr += attrBytes.Length;
                        continue;
                    }

                    if (attrEntry.Value is List<string> attrStrings)
                    {
                        stringAttrs.AddRange(attrStrings);
                        continue;
                    }
                    
                    var attrInfo = attrInfos.FirstOrDefault(x => x.Name == attrEntry.Key);
                    if (attrInfo != null)
                    {
                        curBytesPerAttr += GeneralUtils.GetTypeSize(attrInfo.Type);
                        switch (attrInfo.Type)
                        {
                            case ParamType.UInt8:
                                writer.WriteByte((byte)attrEntry.Value);
                                break;
                            case ParamType.UInt16:
                                writer.WriteUInt16((ushort)attrEntry.Value);
                                break;
                            case ParamType.UInt32:
                                writer.WriteUInt32((uint)attrEntry.Value);
                                break;
                            case ParamType.Int8:
                                writer.WriteByte((byte)attrEntry.Value);
                                break;
                            case ParamType.Int16:
                                writer.WriteInt16((short)attrEntry.Value);
                                break;
                            case ParamType.Int32:
                                writer.WriteInt32((int)attrEntry.Value);
                                break;
                            case ParamType.Float:
                                writer.WriteSingle((float)attrEntry.Value);
                                break;
                            case ParamType.Double:
                                writer.WriteDouble((double)attrEntry.Value);
                                break;
                            case ParamType.String:
                                stringAttrs.Add(GeneralUtils.UnquoteString((string)attrEntry.Value));
                                stringAttrOffsetPositions.Add(writer.Position);
                                writer.Pad(4);
                                break;
                            case ParamType.List:
                                int itemId = attrInfo.List.FindIndex(x => x == (string) attrEntry.Value);
                                if (itemId == -1)
                                {
                                    throw new Exception($"Attribute \"{attrEntry.Key}\" doesn't contain an item called \"{(string)attrEntry.Value}\".");
                                }
                                writer.WriteByte((byte)itemId);
                                break;
                            default:
                                throw new Exception($"Attribute {attrEntry.Key} has an unknown type!");
                        }
                    }
                    else
                    {
                        throw new Exception($"Can't find an attribute called \"{attrEntry.Key}\" inside the MSBP.");
                    }
                }

                if (i == 0)
                {
                    bytesPerAttr = curBytesPerAttr;
                }
                else
                {
                    if (curBytesPerAttr != bytesPerAttr)
                    {
                        throw new Exception($"Message {i} doesn't have the same amount of attributes as message 0.");
                    }
                }
            }

            writer.WriteInt32AtAndReturn(bytesPerAttrPosition, bytesPerAttr);

            for (int i = 0; i < stringAttrs.Count; i++)
            {
                if (stringAttrOffsetPositions.Count > 0)
                {
                    int stringOffset = (int) (writer.Position - startPosition);
                    writer.WriteInt32AtAndReturn(stringAttrOffsetPositions[i], stringOffset);
                }
                writer.WriteTerminatedString(stringAttrs[i], encoding);
            }
            
            CalculateAndSetSectionSize(writer, sizePosition);
            writer.Align(0x10, 0xAB);
        }
    }
    internal class TSY1
    {
        public List<int> StyleIndices { get; set; }

        public TSY1() {}

        public TSY1(FileReader reader, long sectionSize)
        {
            StyleIndices = new();
            
            for (uint i = 0; i < sectionSize / 4; i++)
            {
                StyleIndices.Add(reader.ReadInt32());
            }
        }

        public static void Write(FileWriter writer, Message[] messages)
        {
            writer.WriteString("TSY1", Encoding.ASCII);
            long sizePosition = writer.Position;
            writer.Pad(0xC);

            foreach (var message in messages)
            {
                writer.WriteUInt32((uint)message.StyleId);
            }
            
            CalculateAndSetSectionSize(writer, sizePosition);
            writer.Align(0x10, 0xAB);
        }
    }
    internal class TXT2
    {
        public List<Message> Messages { get; set; }

        public TXT2() {}

        public TXT2(FileReader reader, bool hasATR1, ATR1 atr1, bool hasTSY1, TSY1 tsy1, ParsingOptions options, LBL1 lbl1, bool hasNLI1, NLI1 nli1, string fileName, Encoding encoding, MSBP? msbp = null)
        {
            Messages = new();
            //encoding = Encoding.Unicode;
            long startPosition = reader.Position;
            uint messageCount = reader.ReadUInt32();
            List<uint> stringOffsets = new();

            for (uint i = 0; i < messageCount; i++)
            {
                stringOffsets.Add(reader.ReadUInt32());
            }

            for (uint i = 0; i < messageCount; i++)
            {
                //Console.WriteLine(lbl1.Labels[Messages.Count]);
                Message message = new();
                if (hasATR1)
                {
                    if (i < atr1.AttributeDicts.Count)
                    {
                        message.Attributes = atr1.AttributeDicts[(int)i];
                    }
                    else
                    {
                        message.Attributes = new Dictionary<string, object>();
                        //throw new InvalidDataException("Numbers of ATR1 and TXT2 entries don't match!");
                    }
                }
                if (hasTSY1)
                {
                    if (i < tsy1.StyleIndices.Count)
                    {
                        message.StyleId = tsy1.StyleIndices[(int)i];
                    }
                    else
                    {
                        message.StyleId = -1;
                        //throw new InvalidDataException("Numbers of TSY1 and TXT2 entries don't match!");
                    }
                }
                
                reader.JumpTo(startPosition + stringOffsets[(int)i]);

                bool reachedEnd = false;
                
                long nextStringPosition = i + 1 < messageCount ?
                        startPosition + stringOffsets[(int)i + 1] :
                        FindPosOfEndFileOrAb(reader, startPosition + stringOffsets[(int)i]);
                    
                StringBuilder messageString = new();

                while (!reachedEnd && reader.Position < nextStringPosition)
                {
                    char character;
                    switch (encoding)
                    {
                        case UTF8Encoding:
                            List<byte> charBytes = new();
                            character = '\ufffd';
                            while (character == '\ufffd')
                            {
                                charBytes.Add(reader.ReadByte());
                                character = encoding.GetChars(charBytes.ToArray())[0];
                            }
                            break;
                        case UnicodeEncoding:
                            if (encoding.BodyName == "utf-16BE")
                            {
                                character = encoding.GetChars(BitConverter.GetBytes(reader.ReadInt16()).Reverse().ToArray())[0];
                            }
                            else
                            {
                                character = encoding.GetChars(BitConverter.GetBytes(reader.ReadInt16()))[0];
                            }
                            break;
                        default:
                            throw new InvalidDataException("Unknown encoding.");
                    }
                    string tagOrigin = hasNLI1 ? $"{fileName}@{nli1.Indices[(uint)Messages.Count]}" : $"{fileName}@{lbl1.Labels[Messages.Count]}";
                    List<byte> buffer = new List<byte>();
                    switch (character)
                    {
                        case '\u000e':
                            switch (encoding)
                            {
                                case UTF8Encoding:
                                    buffer.Add(0xE);
                                    break;
                                case UnicodeEncoding:
                                    if (reader.Endianness == Endianness.LittleEndian)
                                    {
                                        buffer.AddRange(new byte[]{0xE, 0x0});
                                    }
                                    else
                                    {
                                        buffer.AddRange(new byte[]{0x0, 0xE});
                                    }
                                    break;
                                default:
                                    throw new InvalidDataException("Unknown encoding.");
                            }
                            ushort tagGroup;
                            ushort tagType;
                            ushort argumentsLength;
                            byte[] rawTagParameters;
                            
                            try
                            {
                                if (reader.Position >= nextStringPosition) throw new InvalidDataException();
                                buffer.AddRange(reader.PeekBytes(2));
                                tagGroup = reader.ReadUInt16();
                                
                                if (reader.Position >= nextStringPosition) throw new InvalidDataException();
                                buffer.AddRange(reader.PeekBytes(2));
                                tagType = reader.ReadUInt16();
                                
                                if (reader.Position >= nextStringPosition) throw new InvalidDataException();
                                buffer.AddRange(reader.PeekBytes(2));
                                argumentsLength = reader.ReadUInt16();
                                
                                if (reader.Position + argumentsLength > nextStringPosition)
                                {
                                    buffer.AddRange(reader.PeekBytes((int)(nextStringPosition - reader.Position)));
                                    throw new InvalidDataException();
                                }
                                rawTagParameters = reader.ReadBytes(argumentsLength);
                            }
                            catch (Exception e) when (e is InvalidDataException)
                            {
                                messageString.Append(GeneralUtils.BytesToEscapeSequences(buffer));
                                break;
                            }
                            
                            Tag tag = new(tagGroup, tagType, rawTagParameters, false);

                            messageString.Append(tag.Stringify(options, tagOrigin, encoding, msbp));
                            break;
                        
                        case '\u000f':
                            switch (encoding)
                            {
                                case UTF8Encoding:
                                    buffer.Add(0xF);
                                    break;
                                case UnicodeEncoding:
                                    if (reader.Endianness == Endianness.LittleEndian)
                                    {
                                        buffer.AddRange(new byte[]{0xF, 0x0});
                                    }
                                    else
                                    {
                                        buffer.AddRange(new byte[]{0x0, 0xF});
                                    }
                                    break;
                                default:
                                    throw new InvalidDataException("Unknown encoding.");
                            }
                            ushort tagEndGroup;
                            ushort tagEndType;

                            try
                            {
                                if (reader.Position >= nextStringPosition) throw new InvalidDataException();
                                buffer.AddRange(reader.PeekBytes(2));
                                tagEndGroup = reader.ReadUInt16();
                                
                                if (reader.Position >= nextStringPosition) throw new InvalidDataException();
                                buffer.AddRange(reader.PeekBytes(2));
                                tagEndType = reader.ReadUInt16();
                            }
                            catch (Exception e) when (e is InvalidDataException)
                            {
                                messageString.Append(GeneralUtils.BytesToEscapeSequences(buffer));
                                break;
                            }

                            Tag tagEnd = new(tagEndGroup, tagEndType, new byte[0], true);

                            messageString.Append(tagEnd.Stringify(options, tagOrigin, encoding, msbp));
                            break;
                        
                        case '\0':
                            reachedEnd = true;
                            break;
                        
                        case '<':
                            messageString.Append('\\');
                            messageString.Append('<');
                            break;
                        
                        case '>':
                            messageString.Append('\\');
                            messageString.Append('>');
                            break;
                        
                        case '\\':
                            messageString.Append('\\');
                            messageString.Append('\\');
                            break;
                        
                        default:
                            messageString.Append(character);
                            break;
                    }
                }

                message.Text = messageString.ToString();
                Messages.Add(message);
                //Console.WriteLine($"{i}: {message.Text}");
            }
        }

        private static long FindPosOfEndFileOrAb(FileReader reader, long startPos)
        {
            long prevPos = reader.Position;
            long pos = 0;
            
            while (!reader.AtEndOfStream)
            {
                if (reader.ReadByte() == 0xAB)
                {
                    if (reader.AtEndOfStream)
                    {
                        pos = reader.Position;
                        break;
                    }
                    else
                    {
                        if (reader.ReadByte() == 0xAB)
                        {
                            pos = reader.Position - 1;
                            break;
                        }
                        else
                        {
                            reader.Position--;
                        }
                    }
                }

                pos = reader.Position;
            }
            
            reader.Position = prevPos;

            return pos;
        }

        public static void Write(FileWriter writer, Message[] messages, Encoding encoding, ParsingOptions options, MSBP? msbp = null)
        {
            writer.WriteString("TXT2", Encoding.ASCII);
            long sizePosition = writer.Position;
            writer.Pad(0xC);

            long startPosition = writer.Position;
            writer.WriteUInt32((uint)messages.Length);
            long offsetsPosition = writer.Position;
            writer.Skip(messages.Length * 4);

            for (int i = 0; i < messages.Length; i++)
            {
                long messagePosition = writer.Position;
                writer.JumpTo(offsetsPosition + i * 4);
                writer.WriteUInt32((uint) (messagePosition - startPosition));
                writer.JumpTo(messagePosition);

                string text = messages[i].Text;
                text = text.Replace("\r\n", "\n");
                
                for (int j = 0; j < text.Length; j++)
                {
                    char c = text[j];
                    switch (c)
                    {
                        case '<':
                            string tagText = "";
                            for (int k = j; k < text.Length; k++)
                            {
                                if (text[k] == '>')
                                {
                                    j = k;
                                    tagText += '>';
                                    break;
                                }
                                tagText += text[k];
                            }

                            try
                            {
                                byte[] tagBytes = Tag.Write(writer, tagText, options, encoding is UTF8Encoding, msbp);
                                if (options.AddLinebreaksAfterPagebreaks && text.Length != j + 1)
                                {
                                    bool isPageBreakTag = encoding is UTF8Encoding
                                        ? (tagBytes[0] == 0xE && tagBytes[1] == 0x0 && tagBytes[3] == 0x4) ||
                                          (tagBytes[0] == 0xE && tagBytes[2] == 0x0 && tagBytes[4] == 0x4)
                                        : (tagBytes[0] == 0xE && tagBytes[2] == 0x0 && tagBytes[4] == 0x4) ||
                                          (tagBytes[1] == 0xE && tagBytes[3] == 0x0 && tagBytes[5] == 0x4);
                                    if (isPageBreakTag)
                                    {
                                        //Console.WriteLine(text[j + 1]);
                                        if (text[j + 1] == '\n')
                                        {
                                            j++;
                                        }
                                        // else
                                        // {
                                        //     Console.WriteLine("Alert");
                                        // }
                                    }
                                }
                            }
                            catch
                            {
                                Console.WriteLine($"Couldn't parse tag {tagText}");
                                writer.WriteString(tagText, encoding);
                            }

                            break;
                        case '\\':
                            if (j + 1 == text.Length)
                            {
                                writer.WriteString(c.ToString(), encoding);
                                break;
                            }

                            char nextChar = text[j + 1];
                            switch (nextChar)
                            {
                                case '<':
                                    writer.WriteString("<", encoding);
                                    j++;
                                    break;
                                case '>':
                                    writer.WriteString(">", encoding);
                                    j++;
                                    break;
                                case 'x':
                                    if (j + 3 < text.Length)
                                    {
                                        string strByte = $"{text[j + 2]}{text[j + 3]}";
                                        byte b = Convert.FromHexString(strByte)[0];
                                        writer.WriteByte(b);
                                        j += 3;
                                    }
                                    else
                                    {
                                        writer.WriteString("\\", encoding);
                                    }
                                    break;
                                case '\\':
                                    writer.WriteString("\\", encoding);
                                    j++;
                                    break;
                                default:
                                    writer.WriteString("\\", encoding);
                                    break;
                            }

                            break;
                        default:
                            writer.WriteString(c.ToString(), encoding);
                            break;
                    }
                }
                writer.WriteString('\0'.ToString(), encoding);
            }
            
            CalculateAndSetSectionSize(writer, sizePosition);
            writer.Align(0x10, 0xAB);
        }
    }
    
    #endregion
}
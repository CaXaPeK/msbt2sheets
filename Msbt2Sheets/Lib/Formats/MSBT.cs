using System.Data;
using System.Text;
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

    public bool UsesAttributeStrings { get; set; }

    public uint BytesPerAttribute { get; set; }
    
    public uint LabelSlotCount { get; set; }

    public List<int> ATO1Numbers = new();

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
                    atr1 = new(reader, sectionSize, this);
                    break;
                case "TSY1":
                    HasTSY1 = true;
                    tsy1 = new(reader, sectionSize);
                    break;
                case "TXT2":
                    txt2 = new(reader, HasATR1, atr1, HasTSY1, tsy1, options, lbl1, fileName, Header.Encoding, msbp);
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
            foreach (var pair in nli1.Indices)
            {
                Messages.Add(pair.Key.ToString(), txt2.Messages[(int)pair.Value]);
            }
        }

        if (HasATO1)
        {
            ATO1Numbers = ato1.Numbers;
        }
    }

    public void PrintAllMessages()
    {
        Console.WriteLine($"{Language} {FileName}.msbt");
        foreach (var Message in Messages)
        {
            Console.WriteLine($"[{Message.Key}] {Message.Value.Text}");
        }
    }

    public List<string> MessagesToStringList()
    {
        List<string> strings = new();
        foreach (var Message in Messages)
        {
            strings.Add($"[{Message.Key}] {Message.Value.Text}");
        }

        return strings;
    }

    public byte[] Compile(MSBP? msbp = null)
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
            ATO1.Write(writer, ATO1Numbers);
            sectionCount++;
        }
        
        if (HasATR1)
        {
            ATR1.Write(writer, Messages.Values.ToArray(), BytesPerAttribute, UsesAttributeStrings);
            sectionCount++;
        }
        
        if (HasTSY1)
        {
            TSY1.Write(writer, Messages.Values.ToArray());
            sectionCount++;
        }
        
        TXT2.Write(writer, Messages.Values.ToArray(), Header.Encoding, msbp);
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
                uint entryId = reader.ReadUInt32();
                uint messageIndex = reader.ReadUInt32();
                Indices.Add(entryId, messageIndex);
            }
        }
    }
    internal class ATO1
    {
        public List<int> Numbers { get; set; }

        public ATO1() {}

        public ATO1(FileReader reader, long sectionSize)
        {
            Numbers = new();
            
            for (long i = 0; i < sectionSize / 4; i++)
            {
                Numbers.Add(reader.ReadInt32());
            }
        }

        public static void Write(FileWriter writer, List<int> numbers)
        {
            writer.WriteString("ATO1", Encoding.ASCII);
            long sizePosition = writer.Position;
            writer.Pad(0xC);

            foreach (var number in numbers)
            {
                writer.WriteInt32(number);
            }
            
            CalculateAndSetSectionSize(writer, sizePosition);
            writer.Align(0x10, 0xAB);
        }
    }
    internal class ATR1
    {
        public List<MessageAttribute> Attributes { get; set; }

        public ATR1() {}

        public ATR1(FileReader reader, long sectionSize, MSBT msbt)
        {
            Attributes = new();
            
            long startPosition = reader.Position;
            uint attributeCount = reader.ReadUInt32();
            uint bytesPerAttribute = reader.ReadUInt32();
            bool hasStrings = sectionSize > 8 + attributeCount * bytesPerAttribute;
            
            msbt.BytesPerAttribute = bytesPerAttribute;
            msbt.UsesAttributeStrings = hasStrings;
            
            List<byte[]> attributeByteData = new();
            for (uint i = 0; i < attributeCount; i++)
            {
                attributeByteData.Add(reader.ReadBytes((int)bytesPerAttribute));
            }

            foreach (byte[] byteData in attributeByteData)
            {
                if (hasStrings)
                {
                    string stringData = reader.ReadTerminatedString(msbt.Header.Encoding);
                    Attributes.Add(new(byteData, stringData));
                    
                    /*if ((BitConverter.IsLittleEndian && reader.Endianness == Endianness.BigEndian) ||
                        (!BitConverter.IsLittleEndian && reader.Endianness == Endianness.LittleEndian))
                    {
                        Array.Reverse(byteData);
                    }
                    
                    uint stringOffset = BitConverter.ToUInt32(byteData[..4]);
                    string stringData = reader.ReadTerminatedStringAt(startPosition + stringOffset);
                
                    Attributes.Add(new(byteData, stringData));*/
                }
                else
                {
                    Attributes.Add(new(byteData));
                }
            }
        }

        public static void Write(FileWriter writer, Message[] messages, uint bytesPerAttribute, bool usesAttributeStrings)
        {
            writer.WriteString("ATR1", Encoding.ASCII);
            long sizePosition = writer.Position;
            writer.Pad(0xC);
            
            long startPosition = writer.Position;
            writer.WriteUInt32((uint)messages.Length);
            writer.WriteUInt32(bytesPerAttribute);

            if (usesAttributeStrings)
            {
                long hashTablePosition = writer.Position;
                writer.Skip(messages.Length * bytesPerAttribute);
                for (int i = 0; i < messages.Length; i++)
                {
                    long attributePosition = writer.Position;

                    byte[] attribute = messages[i].Attribute.ByteData;
                    uint stringOffset = (uint)(attributePosition - startPosition);
                    long stringPosition = writer.Position;
                    
                    writer.JumpTo(hashTablePosition + (i * bytesPerAttribute));
                    writer.WriteBytes(attribute);
                    writer.WriteUInt32(stringOffset);
                    
                    long nextAttributePosition = writer.Position;
                    
                    writer.JumpTo(stringPosition);
                    writer.WriteString(messages[i].Attribute.StringData + '\0', Encoding.ASCII);
                    writer.JumpTo(nextAttributePosition);
                }
            }
            else
            {
                foreach (var message in messages)
                {
                    writer.WriteBytes(message.Attribute.ByteData);
                }
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

        public TXT2(FileReader reader, bool hasATR1, ATR1 atr1, bool hasTSY1, TSY1 tsy1, ParsingOptions options, LBL1 lbl1, string fileName, Encoding encoding, MSBP? msbp = null)
        {
            Messages = new();
            
            long startPosition = reader.Position;
            uint messageCount = reader.ReadUInt32();

            for (uint i = 0; i < messageCount; i++)
            {
                Message message = new();
                if (hasATR1)
                {
                    if (i < atr1.Attributes.Count)
                    {
                        message.Attribute = atr1.Attributes[(int)i];
                    }
                    else
                    {
                        throw new InvalidDataException("Numbers of ATR1 and TXT2 entries don't match!");
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

                uint stringOffset = reader.ReadUInt32();
                long nextOffsetPosition = reader.Position;
                
                reader.JumpTo(startPosition + stringOffset);

                bool reachedEnd = false;
                StringBuilder messageString = new();

                while (!reachedEnd)
                {
                    short character = reader.ReadInt16();
                    string tagOrigin = $"{fileName}@{lbl1.Labels[Messages.Count]}";
                    switch (character)
                    {
                        case 0x0E:
                            ushort tagGroup = reader.ReadUInt16();
                            ushort tagType = reader.ReadUInt16();
                            ushort argumentsLength = reader.ReadUInt16();
                            byte[] rawTagParameters = reader.ReadBytes(argumentsLength);

                            Tag tag = new(tagGroup, tagType, rawTagParameters, false);

                            messageString.Append(tag.Stringify(options, tagOrigin, encoding, msbp));
                            break;
                        
                        case 0x0F:
                            ushort tagEndGroup = reader.ReadUInt16();
                            ushort tagEndType = reader.ReadUInt16();

                            Tag tagEnd = new(tagEndGroup, tagEndType, new byte[0], true);

                            messageString.Append(tagEnd.Stringify(options, tagOrigin, encoding, msbp));
                            break;
                        
                        case 0x00:
                            reachedEnd = true;
                            break;
                        
                        case 0x3C:
                            messageString.Append('\\');
                            messageString.Append('<');
                            break;
                        
                        case 0x3E:
                            messageString.Append('\\');
                            messageString.Append('>');
                            break;
                        
                        case 0x5C:
                            messageString.Append('\\');
                            messageString.Append('\\');
                            break;
                        
                        default:
                            messageString.Append(encoding.GetString(BitConverter.GetBytes(character)));
                            break;
                    }
                }

                message.Text = messageString.ToString();
                Messages.Add(message);
                reader.JumpTo(nextOffsetPosition);
            }
        }

        public static void Write(FileWriter writer, Message[] messages, Encoding encoding, MSBP? msbp = null)
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
                            Tag.Write(writer, tagText, msbp);
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
using System.Text;
using Msbt2Sheets.Lib.Utils;

namespace Msbt2Sheets.Lib.Formats.FileComponents;

public class Header
{
    public FileType FileType { get; set; }
    public Endianness Endianness { get; set; }
    public EncodingType EncodingType { get; set; }
    public byte Version { get; set; }
    public ushort SectionCount = 0;
    public uint FileSize = 0;

    public Encoding Encoding
    {
        get
        {
            if (Endianness == Endianness.BigEndian && EncodingType == EncodingType.UTF16)
            {
                return Encoding.BigEndianUnicode;
            }
            
            switch (EncodingType)
            {
                case EncodingType.UTF8:
                    return Encoding.UTF8;
                case EncodingType.UTF16:
                    return Encoding.Unicode;
                case EncodingType.UTF32:
                    return Encoding.UTF32;
                default:
                    return Encoding.Unicode;
            }
        }
    }

    public Header() {}

    public Header(FileReader reader)
    {
        string magic = reader.ReadString(8, Encoding.ASCII);
        switch (magic)
        {
            case "MsgStdBn":
                FileType = FileType.MSBT;
                break;
            case "MsgPrjBn":
                FileType = FileType.MSBP;
                break;
        }

        Endianness = reader.ReadUInt16() == 0xFFFE ? Endianness.LittleEndian : Endianness.BigEndian;
        reader.Endianness = Endianness;
        
        reader.Skip(2);

        EncodingType = (EncodingType)reader.ReadByte();
        Version = reader.ReadByte();
        SectionCount = reader.ReadUInt16();
        
        reader.Skip(2);

        FileSize = reader.ReadUInt32();
        
        reader.Skip(0xA);
    }

    public void Write(FileWriter writer)
    {
        string magic;
        switch (FileType)
        {
            case FileType.MSBT:
                magic = "MsgStdBn";
                break;
            case FileType.MSBP:
                magic = "MsgPrjBn";
                break;
            default:
                throw new InvalidDataException("Unknown file magic.");
        }
        writer.WriteString(magic, Encoding.ASCII);
        writer.WriteUInt16(0xFEFF);
        writer.Pad(2);
        writer.WriteByte((byte)EncodingType);
        writer.WriteByte(Version);
        writer.WriteUInt16(SectionCount);
        writer.Pad(2);
        writer.WriteUInt32(FileSize);
        writer.Pad(0xA);
    }
}
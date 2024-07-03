namespace Msbt2Sheets.Lib.Formats.FileComponents;

public enum FileType
{
    MSBT,
    MSBP
}

public enum Endianness
{
    BigEndian,
    LittleEndian
}
public enum EncodingType
{
    UTF8,
    UTF16,
    UTF32
}

public enum ParamType : byte
{
    UInt8,
    UInt16,
    UInt32,
    Int8,
    Int16,
    Int32,
    Float,
    Double,
    String,
    List
}
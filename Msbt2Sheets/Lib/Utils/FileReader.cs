﻿using System.Text;
using Msbt2Sheets.Lib.Formats.FileComponents;

namespace Msbt2Sheets.Lib.Utils;

public class FileReader : IDisposable
{
    private readonly BinaryReader _reader;
    private bool _disposed;
    
    public void Dispose()
    {
        if (_disposed) return;
        _reader.Dispose();
        _disposed = true;
    }
    
    public FileReader(Stream fileStream, bool leaveOpen = false)
    {
        _reader = new BinaryReader(fileStream, Encoding.UTF8, leaveOpen);
        Position = 0;
        Endianness = Endianness.BigEndian;
    }
    
    public long Position
    {
        get => _reader.BaseStream.Position;
        set => _reader.BaseStream.Position = value;
    }

    public long Length
    {
        get => _reader.BaseStream.Length;
    }

    public bool AtEndOfStream
    {
        get => Position == Length;
    }
    
    public Endianness Endianness { get; set; }

    public void Skip(int count)
    {
        Position += count;
    }

    public void JumpTo(long position)
    {
        Position = position;
    }

    public void Align(int alignment)
    {
        if (Position % alignment != 0)
        {
            Position = alignment * (Position / alignment + 1);
        }
    }
    
    #region reading

    public byte[] ReadBytes(int length)
    {
        return ReadBytes(length, length);
    }
    
    public byte[] PeekBytes(int length)
    {
        long startPos = Position;
        byte[] bytes = ReadBytes(length, length);
        Position = startPos;
        return bytes;
    }
    
    public byte ReadByte(int length = 1)
    {
        byte[] bytes = ReadBytes(length, 1);
        return bytes[0];
    }
    
    public byte ReadByteAt(long position, int length = 1)
    {
        Position = position;
        return ReadByte(length);
    }
    
    public sbyte ReadSByte(int length = 1)
    {
        byte[] bytes = ReadBytes(length, 1);
        return (sbyte)bytes[0];
    }
    
    public sbyte ReadSByteAt(long position, int length = 1)
    {
        Position = position;
        return ReadSByte(length);
    }
    
    public string ReadString(int length)
    {
        return ReadString(length, Encoding.UTF8);
    }
    
    public string ReadStringAt(long position, int length)
    {
        return ReadStringAt(position, length, Encoding.UTF8);
    }
    
    public string ReadString(int length, Encoding encoding)
    {
        byte[] bytes = ReadBytes(length, 0);
        return encoding.GetString(bytes).TrimEnd('\0');
    }
    
    public string ReadStringAt(long position, int length, Encoding encoding)
    {
        Position = position;
        return ReadString(length, encoding);
    }
    
    public string ReadTerminatedString(int maxLength = -1)
    {
        return ReadTerminatedString(Encoding.UTF8, maxLength);
    }
    
    public string ReadTerminatedStringAt(long position, int maxLength = -1)
    {
        return ReadTerminatedStringAt(position, Encoding.UTF8, maxLength);
    }
    
    public string PeekTerminatedString(long position, Encoding encoding, int maxLength = -1)
    {
        long prevPosition = Position;
        string result = ReadTerminatedStringAt(position, encoding, maxLength);
        JumpTo(prevPosition);
        return result;
    }

    public string ReadTerminatedString(Encoding encoding, int maxLength = -1)
    {
        List<byte> bytes = new(maxLength > 0 ? maxLength : 256);
        int charLength = encoding.GetByteCount("\0");
        while (true)
        {
            byte[] charBytes = _reader.ReadBytes(charLength);
            if (encoding.GetString(charBytes) == "\0")
            {
                return encoding.GetString(bytes.ToArray());
            }
            bytes.AddRange(charBytes);
        }
    }
    
    public string ReadTerminatedStringAt(long position, Encoding encoding, int maxLength = -1)
    {
        Position = position;
        return ReadTerminatedString(encoding, maxLength);
    }
    
    public short ReadInt16(int length = 2)
    {
        byte[] bytes = ReadBytes(length, 2, Endianness == Endianness.BigEndian);
        return BitConverter.ToInt16(bytes, 0);
    }
    public short ReadInt16At(long position, int length = 2)
    {
        Position = position;
        return ReadInt16(length);
    }
    
    public ushort ReadUInt16(int length = 2)
    {
        byte[] bytes = ReadBytes(length, 2, Endianness == Endianness.BigEndian);
        return BitConverter.ToUInt16(bytes, 0);
    }
    public ushort ReadUInt16At(long position, int length = 2)
    {
        Position = position;
        return ReadUInt16(length);
    }
    
    public int ReadInt32(int length = 4)
    {
        byte[] bytes = ReadBytes(length, 4, Endianness == Endianness.BigEndian);
        return BitConverter.ToInt32(bytes, 0);
    }
    public int ReadInt32At(long position, int length = 4)
    {
        Position = position;
        return ReadInt32(length);
    }
    
    public uint ReadUInt32(int length = 4)
    {
        byte[] bytes = ReadBytes(length, 4, Endianness == Endianness.BigEndian);
        return BitConverter.ToUInt32(bytes, 0);
    }
    public uint ReadUInt32At(long position, int length = 4)
    {
        Position = position;
        return ReadUInt32(length);
    }
    
    public float ReadSingle(int length = 4)
    {
        byte[] bytes = ReadBytes(length, 4, Endianness == Endianness.BigEndian);
        return BitConverter.ToSingle(bytes, 0);
    }
    public float ReadSingleAt(long position, int length = 4)
    {
        Position = position;
        return ReadSingle(length);
    }
    
    public double ReadDouble(int length = 8)
    {
        byte[] bytes = ReadBytes(length, 8, Endianness == Endianness.BigEndian);
        return BitConverter.ToSingle(bytes, 0);
    }
    public double ReadDoubleAt(long position, int length = 8)
    {
        Position = position;
        return ReadDouble(length);
    }

    #endregion

    #region private reading
    private byte[] ReadBytes(int length, int padding, bool reversed = false)
    {
        if (length <= 0) return Array.Empty<byte>();

        var bytes = new byte[length > padding ? length : padding];
        _ = _reader.Read(bytes, reversed ? bytes.Length - length : 0, length);

        if (reversed) Array.Reverse(bytes);
        return bytes;
    }

    #endregion
}
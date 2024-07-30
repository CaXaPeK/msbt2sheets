using System.Buffers.Binary;
using System.Text;
using Msbt2Sheets.Lib.Formats.FileComponents;

namespace Msbt2Sheets.Lib.Utils;

public class FileWriter
{
    private readonly BinaryWriter _writer;
    private bool _disposed;
    
    public void Dispose()
    {
        if (_disposed) return;
        _writer.Dispose();
        _disposed = true;
    }
    
    public FileWriter(Stream fileStream, bool leaveOpen = false)
    {
        _writer = new BinaryWriter(fileStream, Encoding.UTF8, leaveOpen);
        Position = 0;
        Endianness = Endianness.BigEndian;
    }
    
    public long Position
    {
        get => _writer.BaseStream.Position;
        set => _writer.BaseStream.Position = value;
    }

    public byte[] Debug
    {
        get => ((MemoryStream) _writer.BaseStream).ToArray();
    }
    
    public Stream BaseStream => _writer.BaseStream;
    
    public Endianness Endianness { get; set; }

    #region writing

    public void Pad(int count)
    {
        _writer.Write(new byte[count]);
    }
    
    public void JumpTo(long position)
    {
        Position = position;
    }
    
    public void Skip(long length)
    {
        Position += length;
    }
    
    public void Align(int alignment, byte value)
    {
        int length = 0;
        if (Position % alignment != 0)
        {
            long newPosition = alignment * (Position / alignment + 1);
            length = (int)(newPosition - Position);
        }

        byte[] alignBytes = new byte[length];
        Array.Fill(alignBytes, value);
        _writer.Write(alignBytes);
    }

    public void WriteBytes(byte[] data)
    {
        _writer.Write(data);
    }
    
    public void WriteByte(byte value)
    {
        _writer.Write(value);
    }
    
    public void WriteInt16(short value)
    {
        byte[] buffer = new byte[sizeof(ushort)];
        if (Endianness == Endianness.LittleEndian)
        {
            BinaryPrimitives.WriteInt16LittleEndian(buffer, value);
        }
        else
        {
            BinaryPrimitives.WriteInt16BigEndian(buffer, value);
        }
        _writer.Write(buffer);
    }

    public void WriteUInt16(ushort value)
    {
        byte[] buffer = new byte[sizeof(ushort)];
        if (Endianness == Endianness.LittleEndian)
        {
            BinaryPrimitives.WriteUInt16LittleEndian(buffer, value);
        }
        else
        {
            BinaryPrimitives.WriteUInt16BigEndian(buffer, value);
        }
        _writer.Write(buffer);
    }
    
    public void WriteUInt32(uint value)
    {
        byte[] buffer = new byte[sizeof(uint)];
        if (Endianness == Endianness.LittleEndian)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(buffer, value);
        }
        else
        {
            BinaryPrimitives.WriteUInt32BigEndian(buffer, value);
        }
        _writer.Write(buffer);
    }
    
    public void WriteInt32(int value)
    {
        byte[] buffer = new byte[sizeof(int)];
        if (Endianness == Endianness.LittleEndian)
        {
            BinaryPrimitives.WriteInt32LittleEndian(buffer, value);
        }
        else
        {
            BinaryPrimitives.WriteInt32BigEndian(buffer, value);
        }
        _writer.Write(buffer);
    }

    public void WriteString(string data, Encoding encoding)
    {
        _writer.Write(encoding.GetBytes(data));
    }

    #endregion
}
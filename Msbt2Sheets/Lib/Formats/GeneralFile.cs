using Msbt2Sheets.Lib.Utils;

namespace Msbt2Sheets.Lib.Formats;

public class GeneralFile
{
    public static uint CalculateLabelCount(FileReader reader)
    {
        long startPosition = reader.Position;
        uint hashTableSlotCount = reader.ReadUInt32();
        uint labelCount = 0;
        for (int i = 0; i < hashTableSlotCount; i++)
        {
            labelCount += reader.ReadUInt32();
            reader.Skip(4);
        }

        reader.Position = startPosition;
        return labelCount;
    }

    public static List<string> ReadLabels(FileReader reader)
    {
        List<string> labels = new(new string[CalculateLabelCount(reader)]);

        long startPosition = reader.Position;
        uint hashTableSlotCount = reader.ReadUInt32();
        for (uint i = 0; i < hashTableSlotCount; i++)
        {
            uint labelCount = reader.ReadUInt32();
            uint labelOffset = reader.ReadUInt32();
            long nextSlotPosition = reader.Position;

            reader.JumpTo(startPosition + labelOffset);
            for (uint j = 0; j < labelCount; j++)
            {
                byte length = reader.ReadByte();
                string labelString = reader.ReadString(length);
                labels[reader.ReadInt32()] = labelString;
            }

            reader.JumpTo(nextSlotPosition);
        }

        return labels;
    }

    public static uint CalculateLabelHash(string label, uint slotCount)
    {
        uint hash = 0;
        foreach (char c in label)
        {
            hash *= 0x492;
            hash += c;
        }

        return (hash & 0xFFFFFFFF) % slotCount;
    }

    public static void CalculateAndSetSectionSize(FileWriter writer, long sizePosition)
    {
        long startPosition = sizePosition + 0xC;
        long endPosition = writer.Position;
        uint sectionSize = (uint)(endPosition - startPosition);
        writer.JumpTo(sizePosition);
        writer.WriteUInt32(sectionSize);
        writer.JumpTo(endPosition);
    }

    public static void CalculateAndSetSectionCountAndFileSize(FileWriter writer, ushort newSectionCount, uint newFileSize)
    {
        writer.JumpTo(0xE);
        writer.WriteUInt16(newSectionCount);
        writer.JumpTo(0x12);
        writer.WriteUInt32(newFileSize);
    }
}
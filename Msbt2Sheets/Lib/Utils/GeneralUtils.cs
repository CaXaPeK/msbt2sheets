using Msbt2Sheets.Lib.Formats;
using Msbt2Sheets.Lib.Formats.FileComponents;

namespace Msbt2Sheets.Lib.Utils;

public class GeneralUtils
{
    public static string ColorToString(System.Drawing.Color color)
    {
        return ($"#{BitConverter.ToString(new byte[]{color.R})}" +
                $"{BitConverter.ToString(new byte[]{color.G})}" +
                $"{BitConverter.ToString(new byte[]{color.B})}" +
                $"{BitConverter.ToString(new byte[]{color.A})}").ToLower();
    }
    
    public static string GetColorNameFromId(int index, MSBP msbp)
    {
        uint i = 0;
        foreach (var color in msbp.Colors)
        {
            if (i == index)
            {
                if (msbp.HasCLB1)
                {
                    return color.Key;
                }
                else
                {
                    return ColorToString(color.Value);
                }
            }

            i++;
        }

        return index.ToString();
    }

    public static int GetTypeSize(ParamType type)
    {
        switch (type)
        {
            case ParamType.UInt8:
                return 1;
            case ParamType.UInt16:
                return 2;
            case ParamType.UInt32:
                return 4;
            case ParamType.Int8:
                return 1;
            case ParamType.Int16:
                return 2;
            case ParamType.Int32:
                return 4;
            case ParamType.Float:
                return 4;
            case ParamType.Double:
                return 8;
            case ParamType.String: //this is for attributes
                return 4;
            case ParamType.List:
                return 1;
            default:
                throw new Exception($"Type {type} is not acceptable as an attribute type.");
        }
    }
    
    public static int GetObjectTypeSize(object obj)
    {
        switch (obj)
        {
            case byte:
                return 1;
            case ushort:
                return 2;
            case uint:
                return 4;
            case sbyte:
                return 1;
            case short:
                return 2;
            case int:
                return 4;
            case float:
                return 4;
            case double:
                return 8;
            case string:
                return 4;
            case List<string>:
                return 1;
            default:
                throw new Exception($"Type {obj.GetType()} is not acceptable as an attribute type.");
        }
    }

    public static string QuoteString(string str)
    {
        str = str.Replace("\\", "\\\\");
        str = str.Replace("\"", "\\\"");
        return '"' + str + '"';
    }

    public static string UnquoteString(string str)
    {
        if (!str.StartsWith('\"') || !str.EndsWith('\"'))
        {
            //throw new Exception($"String {str} doesn't have quotes.");
            return str;
        }

        str = str[1..(str.Length - 1)];
        str = str.Replace("\\\"", "\"");
        str = str.Replace("\\\\", "\\");

        return str;
    }

    public static string BytesToEscapeSequences(List<byte> bytes)
    {
        string str = "";
        foreach (var b in bytes)
        {
            str += $"\\x{BitConverter.ToString(new[]{b})}";
        }

        return str;
    }
}
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
}
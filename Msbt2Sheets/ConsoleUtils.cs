namespace Msbt2Sheets;

public class ConsoleUtils
{
    public static void Error(string err)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(err);
        Console.ResetColor();
        Console.WriteLine();
        Exit();
    }

    public static void CheckFile(string path)
    {
        if (!File.Exists(path))
            Error("File not found: " + path);
    }

    public static void CheckDirectory(string path)
    {
        if (!Directory.Exists(path))
            Error("Folder not found: " + path);
    }

    public static void Exit()
    {
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
        Environment.Exit(0);
    }
    
    public static void Exit(string message)
    {
        Console.WriteLine(message);
        Console.ReadKey();
        Environment.Exit(0);
    }

    public static void WriteColored(string str, ConsoleColor highlightColor)
    {
        ConsoleColor baseColor = Console.ForegroundColor;
        string[] parts = str.Split('|');
        for (int i = 0; i < parts.Length; i++)
        {
            string part = parts[i];
            if (i % 2 == 0)
            {
                Console.ForegroundColor = baseColor;
                Console.Write(part);
            }
            else
            {
                Console.ForegroundColor = highlightColor;
                Console.Write(part);
            }
        }
        Console.ForegroundColor = baseColor;
    }

    public static void WriteLineColored(string str, ConsoleColor highlightColor)
    {
        WriteColored(str, highlightColor);
        Console.WriteLine();
    }
}
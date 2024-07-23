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
}
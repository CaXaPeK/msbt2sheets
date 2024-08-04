using System.Diagnostics;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;using Google.Apis.Sheets.v4.Data;
using Msbt2Sheets;
using Msbt2Sheets.Lib.Formats;
using Msbt2Sheets.Lib.Formats.FileComponents;
using Msbt2Sheets.Lib.Utils;
using Msbt2Sheets.Sheets;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

string credPath = Environment.CurrentDirectory + "/data/credentials.txt";
ConsoleUtils.CheckFile(credPath);

string googleClientId = "";
string googleClientSecret = "";
string[] scopes = { SheetsService.Scope.Spreadsheets };

foreach (string s in File.ReadAllLines(credPath))
{
    if (s.StartsWith("clientId="))
        googleClientId = s.Split('=')[1];
    else if (s.StartsWith("clientSecret="))
        googleClientSecret = s.Split('=')[1];
}

UserCredential credential = GoogleAuth.Login(googleClientId, googleClientSecret, scopes);
GoogleSheetsManager sheetsManager = new GoogleSheetsManager(credential);

static Dictionary<string, string> ParseOptions()
{
    Dictionary<string, string> dict = new();
    
    string optionsPath = Environment.CurrentDirectory + "/data/presets/";
    string[] presets = Directory.GetFiles(optionsPath);
    if (presets.Length == 0)
    {
        return dict;
    }
    else
    {
        Console.WriteLine($"{presets.Length} preset{(presets.Length == 1 ? "" : "s")} found. Which do you want to use?\n\n0: Don't use");
        for (int i = 0; i < presets.Length; i++)
        {
            Console.WriteLine($"{i + 1}: {Path.GetFileNameWithoutExtension(presets[i])}");
        }

        var presetId = Convert.ToInt32(Console.ReadLine()) - 1;
        if (presetId >= 0 && presetId < presets.Length)
        {
            foreach (string s in File.ReadAllLines(presets[presetId]))
            {
                if (s.Contains('='))
                {
                    dict.Add(s.Split('=')[0], s.Split('=')[1]);
                }
            }
        }
    }
    
    return dict;
}

while (true)
{
    Dictionary<string, string> fileOptions = ParseOptions();
    
    Console.Clear();
    Console.WriteLine("Welcome to Msbt2Sheets! What do you want to create?\n\n1 - Spreadsheet\n2 - MSBT folders");
    string mode;
    if (fileOptions.ContainsKey("mode"))
    {
        mode = fileOptions["mode"];
    }
    else
    {
        mode = Console.ReadLine();
    }
    switch (mode)
    {
        case "1":
            MsbtToSheets.Create(sheetsManager, fileOptions);
            ConsoleUtils.Exit();
            break;
        case "2":
            SheetsToMsbt.Create(sheetsManager, fileOptions);
            ConsoleUtils.Exit();
            break;
    }
}
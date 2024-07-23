using Google.Apis.Sheets.v4.Data;
using Msbt2Sheets.Lib.Formats;

namespace Msbt2Sheets.Sheets;

public class SheetsToMsbt
{
    public static void Create(GoogleSheetsManager sheetsManager)
    {
        Console.Clear();
        ConsoleUtils.WriteLineColored("Enter your spreadsheet's ID.\n(It's in the link: https://docs.google.com/spreadsheets/d/|1pRFVKt4fNnWHKf8kIpSk0qmu7u-EdHEUGwkTP9Kzq3A|/edit)", ConsoleColor.Cyan);
        string spreadsheetId = Console.ReadLine();
        
        Console.WriteLine("Loading metadata from the spreadsheet...");
        Spreadsheet spreadsheet = sheetsManager.GetSpreadSheet(spreadsheetId);
        
        List<string> requestRanges = new List<string>();
        foreach (Sheet sheet in spreadsheet.Sheets)
        {
            requestRanges.Add($"{sheet.Properties.Title}!A:ZZZ");
        }
        
        Console.WriteLine("Loading cell data from the spreadsheet...");
        BatchGetValuesResponse valueRanges = sheetsManager.GetMultipleValues(spreadsheetId, requestRanges.ToArray());
        var spreadsheetValues = valueRanges.ValueRanges.ToList();

        ParsingOptions options = ObtainOptions(spreadsheet, spreadsheetValues);
        
        ConsoleUtils.Exit();
    }

    static ParsingOptions ObtainOptions(Spreadsheet spreadsheet, List<ValueRange> valueRanges)
    {
        int sheetId = IndexOfSheetByName(spreadsheet, "#Settings");
        ParsingOptions options = new();
        
        foreach (var row in valueRanges[sheetId].Values)
        {
            if ((string) row[0] == "Add linebreaks after pagebreaks")
            {
                options.AddLinebreaksAfterPagebreaks = (string) row[1] == "TRUE";
            }
            if ((string) row[0] == "Color identification")
            {
                options.ColorIdentification = (string) row[1];
            }
        }

        return options;
    }

    static int IndexOfSheetByName(Spreadsheet spreadsheet, string sheetName)
    {
        int sheetId = -1;
        for (int i = 0; i < spreadsheet.Sheets.Count; i++)
        {
            if (spreadsheet.Sheets[i].Properties.Title == sheetName)
            {
                sheetId = i;
                break;
            }
        }

        return sheetId;
    }
}
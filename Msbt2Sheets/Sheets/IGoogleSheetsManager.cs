using Google.Apis.Sheets.v4.Data;

namespace Msbt2Sheets.Sheets;

public interface IGoogleSheetsManager
{
    Spreadsheet CreateSpreadsheet(Spreadsheet spreadsheet);
    
    Spreadsheet GetSpreadSheet(string googleSpreadsheetIdentifier);

    ValueRange GetSingleValue(string googleSpreadsheetIdentifier, string valueRange);

    BatchGetValuesResponse GetMultipleValues(string googleSpreadsheetIdentifier, string[] ranges);
}
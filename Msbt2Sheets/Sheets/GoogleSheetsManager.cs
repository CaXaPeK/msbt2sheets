using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace Msbt2Sheets.Sheets;

public class GoogleSheetsManager : IGoogleSheetsManager
    {
        private readonly UserCredential _credential;

        public GoogleSheetsManager(UserCredential credential)
        {
            _credential = credential;
        }

        public Spreadsheet CreateSpreadsheet(Spreadsheet spreadsheet)
        {
            if (string.IsNullOrEmpty(spreadsheet.Properties.Title))
                throw new ArgumentNullException(nameof(spreadsheet.Properties.Title));

            using (var sheetsService = new SheetsService(new BaseClientService.Initializer() {HttpClientInitializer = _credential}))
            {
                var documentCreationRequest = sheetsService.Spreadsheets.Create(spreadsheet);

                return documentCreationRequest.Execute();
            }
        }

        public Spreadsheet GetSpreadSheet(string googleSpreadsheetIdentifier)
        {
            if (string.IsNullOrEmpty(googleSpreadsheetIdentifier))
                throw new ArgumentNullException(nameof(googleSpreadsheetIdentifier));

            using (var sheetsService = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = _credential }))
                return sheetsService.Spreadsheets.Get(googleSpreadsheetIdentifier).Execute();
        }

        public ValueRange GetSingleValue(string googleSpreadsheetIdentifier, string valueRange)
        {
            if (string.IsNullOrEmpty(googleSpreadsheetIdentifier))
                throw new ArgumentNullException(nameof(googleSpreadsheetIdentifier));
            if (string.IsNullOrEmpty(valueRange))
                throw new ArgumentNullException(nameof(valueRange));

            using (var sheetsService = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = _credential }))
            {
                var getValueRequest = sheetsService.Spreadsheets.Values.Get(googleSpreadsheetIdentifier, valueRange);
                return getValueRequest.Execute();
            }
        }

        public BatchGetValuesResponse GetMultipleValues(string googleSpreadsheetIdentifier, string[] ranges)
        {
            if (string.IsNullOrEmpty(googleSpreadsheetIdentifier))
                throw new ArgumentNullException(nameof(googleSpreadsheetIdentifier));
            if (ranges == null || ranges.Length <= 0)
                throw new ArgumentNullException(nameof(ranges));

            using (var sheetsService = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = _credential }))
            {
                var getValueRequest = sheetsService.Spreadsheets.Values.BatchGet(googleSpreadsheetIdentifier);
                getValueRequest.Ranges = ranges;
                return getValueRequest.Execute();
            }
        }
    }
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

        public void AddSheets(IList<Sheet> sheets, string spreadsheetId)
        {
            using (var sheetsService = new SheetsService(new BaseClientService.Initializer() {HttpClientInitializer = _credential}))
            {
                var requests = new List<Request>();
                
                foreach (var sheet in sheets)
                {
                    requests.Add(new Request
                    {
                        AddSheet = new AddSheetRequest
                        {
                            Properties = new SheetProperties
                            {
                                Title = sheet.Properties.Title,
                                GridProperties = new GridProperties()
                                {
                                    RowCount = sheet.Properties.GridProperties.RowCount,
                                    ColumnCount = sheet.Properties.GridProperties.ColumnCount,
                                    FrozenRowCount = sheet.Properties.GridProperties.FrozenRowCount,
                                    FrozenColumnCount = sheet.Properties.GridProperties.FrozenColumnCount
                                },
                                SheetId = sheet.Properties.SheetId
                            }
                        }
                    });
                }
                
                var batchAddSheetsRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = requests
                };
                
                try
                {
                    sheetsService.Spreadsheets.BatchUpdate(batchAddSheetsRequest, spreadsheetId).Execute();
                }
                catch (Exception e)
                {
                    if (e.Message.Contains("cancelled"))
                    {
                        Console.WriteLine("Creation request cancelled. Retrying...");
                        sheetsService.Spreadsheets.BatchUpdate(batchAddSheetsRequest, spreadsheetId).Execute();
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }
        
        public void FillSheetsWithFormattingAndWidths(IList<Sheet> sheets, string spreadsheetId)
        {
            using var sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = _credential
            });

            var dataToInsert = new List<ValueRange>();
            var formatRequests = new List<Request>();
            var widthRequests = new List<Request>();

            foreach (var sheet in sheets)
            {
                var values = new List<IList<object>>();
                var sheetId = sheet.Properties.SheetId.Value;
                var sheetTitle = sheet.Properties.Title;

                var rowDataList = sheet.Data?.FirstOrDefault()?.RowData ?? new List<RowData>();

                // 💾 Вставка значений + сбор форматирования
                for (int rowIndex = 0; rowIndex < rowDataList.Count; rowIndex++)
                {
                    var rowData = rowDataList[rowIndex];
                    var row = new List<object>();

                    for (int colIndex = 0; colIndex < rowData.Values.Count; colIndex++)
                    {
                        var cell = rowData.Values[colIndex];
                        var value = cell.UserEnteredValue?.StringValue ?? "";
                        row.Add(value);

                        if (cell.UserEnteredFormat != null)
                        {
                            formatRequests.Add(new Request
                            {
                                RepeatCell = new RepeatCellRequest
                                {
                                    Range = new GridRange
                                    {
                                        SheetId = sheetId,
                                        StartRowIndex = rowIndex,
                                        EndRowIndex = rowIndex + 1,
                                        StartColumnIndex = colIndex,
                                        EndColumnIndex = colIndex + 1
                                    },
                                    Cell = new CellData
                                    {
                                        UserEnteredFormat = cell.UserEnteredFormat
                                    },
                                    Fields = "userEnteredFormat"
                                }
                            });
                        }
                    }

                    values.Add(row);
                }

                // 📐 Ширина столбцов (если есть информация)
                if (sheet.Data != null && sheet.Data.Count > 0 && sheet.Data[0].ColumnMetadata != null)
                {
                    for (int col = 0; col < sheet.Data[0].ColumnMetadata.Count; col++)
                    {
                        var column = sheet.Data[0].ColumnMetadata[col];
                        int? pixelSize = column.PixelSize;
                        if (pixelSize != null)
                        {
                            widthRequests.Add(new Request
                            {
                                UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
                                {
                                    Range = new DimensionRange
                                    {
                                        SheetId = sheetId,
                                        Dimension = "COLUMNS",
                                        StartIndex = col,
                                        EndIndex = col + 1
                                    },
                                    Properties = new DimensionProperties
                                    {
                                        PixelSize = column.PixelSize
                                    },
                                    Fields = "pixelSize"
                                }
                            });
                        }
                    }
                }

                dataToInsert.Add(new ValueRange
                {
                    Range = $"{sheetTitle}!A1",
                    Values = values
                });
            }

            // 1. ⬇ Вставка значений
            var batchValuesRequest = new BatchUpdateValuesRequest
            {
                ValueInputOption = "USER_ENTERED",
                Data = dataToInsert
            };
            
            try
            {
                sheetsService.Spreadsheets.Values.BatchUpdate(batchValuesRequest, spreadsheetId).Execute();
            }
            catch (Exception e)
            {
                if (e.Message.Contains("cancelled"))
                {
                    Console.WriteLine("Value request cancelled. Retrying...");
                    sheetsService.Spreadsheets.Values.BatchUpdate(batchValuesRequest, spreadsheetId).Execute();
                }
                else
                {
                    throw;
                }
            }

            // 2. 🎨 Вставка форматирования и ширины (в одном запросе)
            var batchFormatRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = formatRequests.Concat(widthRequests).ToList()
            };

            if (batchFormatRequest.Requests.Any())
            {
                try
                {
                    sheetsService.Spreadsheets.BatchUpdate(batchFormatRequest, spreadsheetId).Execute();
                }
                catch (Exception e)
                {
                    if (e.Message.Contains("cancelled"))
                    {
                        Console.WriteLine("Format request cancelled. Retrying...");
                        sheetsService.Spreadsheets.BatchUpdate(batchFormatRequest, spreadsheetId).Execute();
                    }
                    else
                    {
                        throw;
                    }
                }
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
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppAutoClick.Helper
{
    public class GoogleSheetsHelper
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "GoogleSheetsHelper";
        private readonly SheetsService _sheetsService;
        private readonly string _spreadsheetId;
        private readonly string _sheetName;
        private List<GoogleSheetRow> _googleSheetRows;

        public GoogleSheetsHelper(string credentialFileName, string spreadsheetId, string sheetName, List<ExpandoObject> datas)
        {
            var credential = GoogleCredential.FromStream(new FileStream(credentialFileName, FileMode.Open)).CreateScoped(Scopes);

            _sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            _spreadsheetId = spreadsheetId;
            _sheetName = sheetName;
            _googleSheetRows = ReadDatas(datas);
        }

        public void WirteDatas()
        {
            var lastRow = GetLastRow(new GoogleSheetParameters { SheetName = _sheetName });
            AddCells(new GoogleSheetParameters { SheetName = _sheetName, RangeRowStart = lastRow, FirstRowIsHeaders = lastRow <= 0 },
                _googleSheetRows);
        }

        private List<GoogleSheetRow> ReadDatas(List<ExpandoObject> datas)
        {
            var list = new List<GoogleSheetRow>();
            foreach (var item in datas)
            {
                if (datas.First().Equals(item))
                {
                    var header = new List<GoogleSheetCell>();
                    foreach(var i in item)
                    {
                        header.Add(new GoogleSheetCell
                        {
                            IsBold = true,
                            CellValue = i.Key != null ? i.Key.ToString() : string.Empty,
                        });
                    }
                    list.Add(new GoogleSheetRow
                    {
                        Cells = header
                    });
                }
                var row = new List<GoogleSheetCell>();
                foreach (var i in item)
                {
                    row.Add(new GoogleSheetCell
                    {
                        CellValue = i.Value != null ? i.Value.ToString() : string.Empty,
                    });
                }
                list.Add(new GoogleSheetRow
                {
                    Cells = row
                });
            }
            return list;
        }

        public int GetLastRow(GoogleSheetParameters googleSheetParameters)
        {
            try
            {
                googleSheetParameters = MakeGoogleSheetDataRangeColumnsZeroBased(googleSheetParameters);
                var range = $"{googleSheetParameters.SheetName}";
                SpreadsheetsResource.ValuesResource.GetRequest request =
                    _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);

                var response = request.Execute();
                if(response.Values != null)
                    return response.Values.Count;
                else
                    return 0;
            }
            catch
            {
                return 0;
            }
        }

        public List<ExpandoObject> GetDataFromSheet(GoogleSheetParameters googleSheetParameters)
        {
            googleSheetParameters = MakeGoogleSheetDataRangeColumnsZeroBased(googleSheetParameters);
            var range = $"{googleSheetParameters.SheetName}!{GetColumnName(googleSheetParameters.RangeColumnStart)}{googleSheetParameters.RangeRowStart}:{GetColumnName(googleSheetParameters.RangeColumnEnd)}{googleSheetParameters.RangeRowEnd}";

            SpreadsheetsResource.ValuesResource.GetRequest request =
                _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);

            var numberOfColumns = googleSheetParameters.RangeColumnEnd - googleSheetParameters.RangeColumnStart;
            var columnNames = new List<string>();
            var returnValues = new List<ExpandoObject>();

            if (!googleSheetParameters.FirstRowIsHeaders)
            {
                for (var i = 0; i <= numberOfColumns; i++)
                {
                    columnNames.Add($"Column{i}");
                }
            }

            var response = request.Execute();

            int rowCounter = 0;
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    if (googleSheetParameters.FirstRowIsHeaders && rowCounter == 0)
                    {
                        for (var i = 0; i <= numberOfColumns; i++)
                        {
                            columnNames.Add(row[i].ToString());
                        }
                        rowCounter++;
                        continue;
                    }

                    var expando = new ExpandoObject();
                    var expandoDict = expando as IDictionary<String, object>;
                    var columnCounter = 0;
                    foreach (var columnName in columnNames)
                    {
                        expandoDict.Add(columnName, row[columnCounter].ToString());
                        columnCounter++;
                    }
                    returnValues.Add(expando);
                    rowCounter++;
                }
            }

            return returnValues;
        }

        public void AddCells(GoogleSheetParameters googleSheetParameters, List<GoogleSheetRow> rows)
        {
            var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };

            var sheetId = GetSheetId(_sheetsService, _spreadsheetId, googleSheetParameters.SheetName);

            GridCoordinate gc = new GridCoordinate
            {
                ColumnIndex = googleSheetParameters.RangeColumnStart,
                RowIndex = googleSheetParameters.RangeRowStart,
                SheetId = sheetId
            };

            var request = new Request { UpdateCells = new UpdateCellsRequest { Start = gc, Fields = "*" } };

            var listRowData = new List<RowData>();

            foreach (var row in rows)
            {
                var rowData = new RowData();
                var listCellData = new List<CellData>();
                if (rows.First().Equals(row) && !googleSheetParameters.FirstRowIsHeaders)
                    continue;
                foreach (var cell in row.Cells)
                {
                    var cellData = new CellData();
                    var extendedValue = new ExtendedValue { StringValue = cell.CellValue };

                    cellData.UserEnteredValue = extendedValue;
                    var cellFormat = new CellFormat { TextFormat = new TextFormat() };

                    if (cell.IsBold)
                    {
                        cellFormat.TextFormat.Bold = true;
                    }

                    cellFormat.BackgroundColor = new Color { Blue = (float)cell.BackgroundColor.B / 255, Red = (float)cell.BackgroundColor.R / 255, Green = (float)cell.BackgroundColor.G / 255 };

                    cellData.UserEnteredFormat = cellFormat;
                    listCellData.Add(cellData);
                }
                rowData.Values = listCellData;
                listRowData.Add(rowData);
            }
            request.UpdateCells.Rows = listRowData;

            // It's a batch request so you can create more than one request and send them all in one batch. Just use reqs.Requests.Add() to add additional requests for the same spreadsheet
            requests.Requests.Add(request);

            _sheetsService.Spreadsheets.BatchUpdate(requests, _spreadsheetId).Execute();
        }

        private string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];
            return value;
        }

        private GoogleSheetParameters MakeGoogleSheetDataRangeColumnsZeroBased(GoogleSheetParameters googleSheetParameters)
        {
            googleSheetParameters.RangeColumnStart = googleSheetParameters.RangeColumnStart - 1;
            googleSheetParameters.RangeColumnEnd = googleSheetParameters.RangeColumnEnd - 1;
            return googleSheetParameters;
        }

        private int GetSheetId(SheetsService service, string spreadSheetId, string spreadSheetName)
        {
            var spreadsheet = service.Spreadsheets.Get(spreadSheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == spreadSheetName);
            int sheetId = (int)sheet.Properties.SheetId;
            return sheetId;
        }
    }

    public class GoogleSheetCell
    {
        public string CellValue { get; set; }
        public bool IsBold { get; set; }
        public System.Drawing.Color BackgroundColor { get; set; } = System.Drawing.Color.White;
    }

    public class GoogleSheetRow
    {
        public GoogleSheetRow() => Cells = new List<GoogleSheetCell>();

        public List<GoogleSheetCell> Cells { get; set; }
    }
}

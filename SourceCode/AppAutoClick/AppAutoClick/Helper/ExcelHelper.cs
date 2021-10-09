using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppAutoClick.Helper
{
    public class ExcelHelper
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "GoogleSheetsHelper";
        private readonly SheetsService _sheetsService;
        private readonly string _spreadsheetId;
        private readonly string _fileName;
        private readonly int _lastRow;
        private readonly GoogleSheetParameters _googleSheetParameters;

        public ExcelHelper(string credentialFileName, string spreadsheetId, string sheetName, string fileName)
        {
            try
            {
                var credential = GoogleCredential.FromStream(new FileStream(credentialFileName, FileMode.Open)).CreateScoped(Scopes);

                _sheetsService = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                _spreadsheetId = spreadsheetId;
                _fileName = fileName;

                _googleSheetParameters = new GoogleSheetParameters();
                _googleSheetParameters.SheetName = sheetName;

                _lastRow = GetLastRow(_googleSheetParameters);
                _googleSheetParameters.RangeRowStart = _lastRow > 0 ? 1 : 0;
                _googleSheetParameters.FirstRowIsHeaders = _lastRow <= 0;
            }
            catch (Exception ex)
            {
                LoggingHelper.Write("ExcelHelper: " + ex.Message);
                throw new InvalidOperationException("Lỗi tạo file google sheet");
            }
        }

        public int GetLastRow(GoogleSheetParameters googleSheetParameters)
        {
            try
            {
                googleSheetParameters = MakeGoogleSheetDataRangeColumnsZeroBased(googleSheetParameters);
                var range = $"{googleSheetParameters.SheetName}";
                SpreadsheetsResource.ValuesResource.GetRequest request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);

                var response = request.Execute();
                if (response.Values != null)
                    return response.Values.Count;
                else
                    return 0;
            }
            catch(Exception ex)
            {
                LoggingHelper.Write("GetLastRow: " + ex.Message);
                throw new InvalidOperationException("Lỗi đọc dữ liệu file google sheet");
            }
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
        private int GetSheetId(SheetsService service, string spreadSheetId, string spreadSheetName)
        {
            try
            {
                var spreadsheet = service.Spreadsheets.Get(spreadSheetId).Execute();
                var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == spreadSheetName);
                int sheetId = (int)sheet.Properties.SheetId;
                return sheetId;
            }
            catch (Exception ex)
            {
                LoggingHelper.Write("GetSheetId: " + ex.Message);
                throw new InvalidOperationException("Lỗi sử dụng file google sheet");
            }
        }

        private GoogleSheetParameters MakeGoogleSheetDataRangeColumnsZeroBased(GoogleSheetParameters googleSheetParameters)
        {
            try
            {
                googleSheetParameters.RangeColumnStart = googleSheetParameters.RangeColumnStart;
                googleSheetParameters.RangeColumnEnd = googleSheetParameters.RangeColumnEnd;
                return googleSheetParameters;
            }
            catch (Exception ex)
            {
                LoggingHelper.Write("MakeGoogleSheetDataRangeColumnsZeroBased: " + ex.Message);
                throw new InvalidOperationException("Lỗi sử dụng file google sheet");
            }
        }

        public void DeleteRowGoogleSheet(int sheetId)
        {
            try
            {
                if (_lastRow > 1)
                {
                    Request request = new Request()
                    {
                        DeleteDimension = new DeleteDimensionRequest()
                        {
                            Range = new DimensionRange()
                            {
                                SheetId = sheetId,
                                Dimension = "ROWS",
                                StartIndex = 1,
                                EndIndex = _lastRow
                            }
                        }
                    };

                    List<Request> requests = new List<Request>();
                    requests.Add(request);

                    BatchUpdateSpreadsheetRequest deleteRequest = new BatchUpdateSpreadsheetRequest();
                    deleteRequest.Requests = requests;

                    _sheetsService.Spreadsheets.BatchUpdate(deleteRequest, _spreadsheetId).Execute();
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Write("DeleteRowGoogleSheet: " + ex.Message);
                throw new InvalidOperationException("Lỗi xóa dòng google sheet");
            }
        }

        public void ReadFileExcel()
        {
            try
            {
                var requests = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };

                var sheetId = GetSheetId(_sheetsService, _spreadsheetId, _googleSheetParameters.SheetName);
                DeleteRowGoogleSheet(sheetId);

                GridCoordinate gc = new GridCoordinate
                {
                    ColumnIndex = _googleSheetParameters.RangeColumnStart,
                    RowIndex = _googleSheetParameters.RangeRowStart,
                    SheetId = sheetId
                };

                var request = new Request { UpdateCells = new UpdateCellsRequest { Start = gc, Fields = "*" } };

                var listRowData = new List<RowData>();
                IWorkbook workbook = new XSSFWorkbook();
                FileStream fs = new FileStream(_fileName, FileMode.Open, FileAccess.Read);
                if (_fileName.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                else if (_fileName.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);
                //First sheet
                ISheet sheet = workbook.GetSheetAt(0);
                if (sheet != null)
                {
                    int rowCount = sheet.LastRowNum; // This may not be valid row count.
                                                     // If first row is table head, i starts from 1

                    DimensionRange dr = new DimensionRange
                    {
                        SheetId = sheetId,
                        Dimension = "ROWS",
                        StartIndex = 1,
                        EndIndex = rowCount // adding extra 6000 rows
                    };
                    var requestAddRow = new Request { InsertDimension = new InsertDimensionRequest { Range = dr, InheritFromBefore = false } };
                    requests.Requests.Add(requestAddRow);

                    var headers = new List<string>();
                    for (int i = 0; i <= rowCount; i++)
                    {
                        var rowData = new RowData();
                        var listCellData = new List<CellData>();
                        IRow curRow = sheet.GetRow(i);
                        // Works for consecutive data. Use continue otherwise 
                        if (curRow == null)
                        {
                            continue;
                        }
                        int cellCount = curRow.LastCellNum;
                        if (i == 0 && !_googleSheetParameters.FirstRowIsHeaders)
                            continue;
                        for (int j = 0; j < cellCount; j++)
                        {
                            var cellData = new CellData();
                            var extendedValue = new ExtendedValue();
                            var cell = curRow.GetCell(j);
                            if (cell != null)
                            {
                                switch (cell.CellType)
                                {
                                    case CellType.Blank: extendedValue.StringValue = string.Empty; break;
                                    case CellType.Boolean: extendedValue.BoolValue = cell.BooleanCellValue; break;
                                    case CellType.String: extendedValue.StringValue = cell.StringCellValue; break;
                                    case CellType.Numeric:
                                        if (HSSFDateUtil.IsCellDateFormatted(cell)) { extendedValue.StringValue = cell.DateCellValue.ToString("mm/dd/yyyy"); }
                                        else { extendedValue.NumberValue = cell.NumericCellValue; }
                                        break;
                                    case CellType.Formula:
                                        switch (cell.CachedFormulaResultType)
                                        {
                                            case CellType.Blank: extendedValue.StringValue = string.Empty; break;
                                            case CellType.String: extendedValue.StringValue = cell.StringCellValue; break;
                                            case CellType.Boolean: extendedValue.BoolValue = cell.BooleanCellValue; break;
                                            case CellType.Numeric:
                                                if (HSSFDateUtil.IsCellDateFormatted(cell)) { extendedValue.StringValue = cell.DateCellValue.ToString("mm/dd/yyyy"); }
                                                else { extendedValue.NumberValue = cell.NumericCellValue; }
                                                break;
                                        }
                                        break;
                                    default: extendedValue.StringValue = cell.StringCellValue; break;
                                }
                            }
                            cellData.UserEnteredValue = extendedValue;
                            var cellFormat = new CellFormat { TextFormat = new TextFormat() };
                            var border = new Border();
                            border.Color = new Color { Red = 0, Green = 0, Blue = 0 };
                            border.Width = 1;
                            border.Style = "SOLID";
                            cellFormat.Borders = new Borders { Top = border, Right = border, Bottom = border, Left = border };
                            
                            if (i == 0)
                            {
                                cellFormat.TextFormat.Bold = true;
                            }

                            //cellFormat.BackgroundColor = new Color { Blue = System.Drawing.Color.Blue / 255, Red = (float)cell.BackgroundColor.R / 255, Green = (float)cell.BackgroundColor.G / 255 };

                            cellData.UserEnteredFormat = cellFormat;
                            listCellData.Add(cellData);
                        }
                        rowData.Values = listCellData;
                        listRowData.Add(rowData);
                    }
                    request.UpdateCells.Rows = listRowData;
                    requests.Requests.Add(request);

                    _sheetsService.Spreadsheets.BatchUpdate(requests, _spreadsheetId).Execute();
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Write("ReadFileExcel: " + ex.Message);
                throw new InvalidOperationException("Lỗi đọc và ghi dữ liệu file google sheet");
            }
        }
    }


    public class GoogleSheetParameters
    {
        public int RangeColumnStart { get; set; }
        public int RangeRowStart { get; set; }
        public int RangeColumnEnd { get; set; }
        public int RangeRowEnd { get; set; }
        public string SheetName { get; set; }
        public bool FirstRowIsHeaders { get; set; }
    }
}

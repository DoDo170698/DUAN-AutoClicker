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
    public static class ExcelHelper
    {
        public static List<ExpandoObject> Datas { get; set; }

        public static List<ExpandoObject> ReadFileExcel(string fileName)
        {
            try
            {
                Datas = new List<ExpandoObject>();
                IWorkbook workbook = new XSSFWorkbook();
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);
                //First sheet
                ISheet sheet = workbook.GetSheetAt(0);
                if (sheet != null)
                {
                    int rowCount = sheet.LastRowNum; // This may not be valid row count.
                                                     // If first row is table head, i starts from 1
                    var headers = new List<string>();
                    for (int i = 0; i <= rowCount; i++)
                    {
                        IRow curRow = sheet.GetRow(i);
                        // Works for consecutive data. Use continue otherwise 
                        if (curRow == null)
                        {
                            break;
                        }
                        int cellCount = curRow.LastCellNum;
                        var expando = new ExpandoObject();
                        var expandoDict = expando as IDictionary<String, object>;
                        for (int j = 0; j < cellCount; j++)
                        {
                            if (i == 0)
                                headers.Add(curRow.GetCell(j).StringCellValue.Trim());
                            else
                            {
                                expandoDict.Add(headers[j], curRow.GetCell(j).StringCellValue.Trim());
                            }
                        }
                        if(expando.Any())
                            Datas.Add(expando);
                    }
                }
                return Datas;
            }
            catch (Exception e)
            {
                return Datas;
            }
        }
    }
}

using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelToJson
{
    public class NPOIUtility
    {
        public static DataTable GetDataFromExcel(string excelFilePath)
        {
            //get exist excel file
            XSSFWorkbook excelBook = new XSSFWorkbook(File.Open(excelFilePath, FileMode.Open));
            var sheet = excelBook.GetSheetAt(0);
            var headerRow = sheet.GetRow(0);

            var table = new DataTable();
            //total columns
            int cellCount = headerRow.LastCellNum;

            //get column name
            for (var i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                var columnName = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(columnName);
            }

            //total rows
            int rowCount = sheet.LastRowNum;
            for (var i = (sheet.FirstRowNum + 1); i <= rowCount; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }

                var dataRow = table.NewRow();

                for (var j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        XSSFFormulaEvaluator e = new XSSFFormulaEvaluator(excelBook);//解析带有公式的Excel
                        var cell = e.EvaluateInCell(row.GetCell(j));
                        dataRow[j] = cell.ToString();
                    }
                }

                table.Rows.Add(dataRow);
            }

            return table;
        }
    }
}

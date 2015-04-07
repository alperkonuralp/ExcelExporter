using System;
using System.Data;
using System.Globalization;
using System.Web;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace ExcelExporter
{
    public class ExcelExporter
    {
        public static byte[] GetExcelBytes(string fileName, DataSet dataSet)
        {
            using (var package = new ExcelPackage())
            {

                var sheetIndex = 1;

                foreach (DataTable table in dataSet.Tables)
                {
                    var tableName = table.TableName;
                    if (String.IsNullOrWhiteSpace(tableName)) tableName = "Sayfa " + sheetIndex;
                    var worksheet = package.Workbook.Worksheets.Add(tableName);

                    worksheet.Cells["A1"].LoadFromDataTable(table, true, TableStyles.Medium1);

                    var i = 1;
                    foreach (DataColumn column in table.Columns)
                    {
                        var excelColumn = worksheet.Column(i);
                        excelColumn.BestFit = true;
                        

                        if (column.DataType == typeof(DateTime))
                        {
                            excelColumn.Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern + " " + DateTimeFormatInfo.CurrentInfo.ShortTimePattern;
                        }

                        excelColumn.AutoFit();
                        i++;
                    }

                    //var columnIndex = 1;
                    //foreach (DataColumn column in table.Columns)
                    //{
                    //    var cell = worksheet.Cells[1, columnIndex];
                    //    cell.Value = column.ColumnName;
                    //    cell.Style.Font.Bold = true;

                    //    columnIndex++;
                    //}

                    //var rowIndex = 2;

                    //foreach (DataRow dataRow in table.Rows)
                    //{
                    //    columnIndex = 1;
                    //    foreach (DataColumn column in table.Columns)
                    //    {
                    //        var excelRange = worksheet.Cells[rowIndex, columnIndex];
                    //        if (dataRow.IsNull(column))
                    //        {
                    //            excelRange.Value = "";
                    //        }
                    //        else
                    //        {
                    //            excelRange.Value = dataRow[column];
                    //            excelRange.Style.Numberformat.Format = "";
                    //        }

                    //        columnIndex++;
                    //    }

                    //    rowIndex++;
                    //}

                    sheetIndex++;
                }

                return package.GetAsByteArray();
            }
        }

        public static void SendExcel(string fileName, DataSet dataSet, HttpResponse response)
        {
            var excel = GetExcelBytes(fileName, dataSet);

            response.ContentEncoding = System.Text.Encoding.UTF8;
            response.Charset = "UTF-8";
            response.AddHeader("content-disposition", "attachment;filename=" + fileName);
            response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            response.BinaryWrite(excel);
        }
    }
}

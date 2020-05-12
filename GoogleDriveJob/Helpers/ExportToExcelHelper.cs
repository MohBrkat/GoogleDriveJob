using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeOpenXml;

namespace GoogleDriveJob.Helpers
{
    public class ExportToExcelHelper
    {
        //public static bool IncludeHeaders { get; set; }
        public static byte[] ExportToExcel<T>(List<T> list, string extension, bool includeHeaders)
        {
            if (extension == "xlsx")
            {
                return ExportXlsxReport(list, includeHeaders);
            }
            else if (extension == "xls")
            {
                return ExportXlsReport(list, includeHeaders);
            }
            else if (extension == "csv")
            {
                return ExportCsvReport(list, includeHeaders);
            }
            else
            {
                throw new Exception("FileFormatIsInvalid");
            }
        }

        public static byte[] ExportXlsxReport<T>(List<T> list, bool includeHeaders)
        {
            var stream = new MemoryStream();

            using (ExcelPackage package = new ExcelPackage(stream))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("report");
                int totalRows = list.Count;

                var properties = typeof(T).GetProperties().ToList();
                if (includeHeaders)
                {
                    for (int j = 1; j <= properties.Count; j++)
                    {
                        var headerValue = Regex.Replace(Regex.Replace(properties[j - 1].Name, @"(\P{Ll})(\P{Ll}\p{Ll})", "$1 $2"), @"(\p{Ll})(\P{Ll})", "$1 $2");
                        worksheet.Cells[1, j].Value = headerValue;
                    }
                }

                int i = 0;

                for (int row = 2; row <= totalRows + 1; row++)
                {
                    for (int j = 1; j <= properties.Count; j++)
                    {
                        worksheet.Cells[row, j].Value = list[i].GetType()
                            .GetProperty(properties[j - 1].Name)
                            ?.GetValue(list[i]);

                        if (worksheet.Cells[row, j].Value is DateTime)
                            worksheet.Cells[row, j].Value = ((DateTime)worksheet.Cells[row, j].Value).ToShortDateString();
                    }

                    i++;
                }

                return package.GetAsByteArray();
            }
        }
        public static byte[] ExportCsvReport<T>(List<T> list, bool includeHeaders)
        {
            var stream = new MemoryStream();
            using (ExcelPackage package = new ExcelPackage(stream))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("report");
                int totalRows = list.Count;

                var properties = typeof(T).GetProperties().ToList();
                StringBuilder sb = new StringBuilder();

                //For Headers
                if (includeHeaders)
                {
                    for (int j = 1; j <= properties.Count; j++)
                    {
                        worksheet.Cells[1, j].Value = Regex.Replace(Regex.Replace(properties[j - 1].Name, @"(\P{Ll})(\P{Ll}\p{Ll})", "$1 $2"), @"(\p{Ll})(\P{Ll})", "$1 $2");
                        if (j == properties.Count)
                        {
                            sb.Append(worksheet.Cells[1, j].Value);
                        }
                        else
                        {
                            sb.Append(worksheet.Cells[1, j].Value + ",");

                        }
                    }

                    sb.AppendLine();
                }


                int i = 0;

                for (int row = 1; row <= totalRows; row++)
                {
                    for (int j = 1; j <= properties.Count; j++)
                    {
                        worksheet.Cells[row, j].Value = list[i].GetType()
                            .GetProperty(properties[j - 1].Name)
                            ?.GetValue(list[i]);

                        if (worksheet.Cells[row, j].Value is DateTime)
                            worksheet.Cells[row, j].Value = ((DateTime)worksheet.Cells[row, j].Value).ToShortDateString();

                        if (j == properties.Count)
                        {
                            sb.Append(worksheet.Cells[row, j].Value);
                        }
                        else
                        {
                            sb.Append(worksheet.Cells[row, j].Value + ",");
                        }
                    }
                    sb.AppendLine();
                    i++;
                }

                return Encoding.ASCII.GetBytes(sb.ToString());
            }
        }
        public static byte[] ExportXlsReport<T>(List<T> list, bool includeHeaders)
        {
            IWorkbook workbook;
            workbook = new HSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet 1");
            IRow row1 = sheet1.CreateRow(0);

            var properties = typeof(T).GetProperties().ToList();

            if (includeHeaders)
            {
                for (int j = 0; j < properties.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(Regex.Replace(Regex.Replace(properties[j].Name, @"(\P{Ll})(\P{Ll}\p{Ll})", "$1 $2"), @"(\p{Ll})(\P{Ll})", "$1 $2"));
                }
            }

            for (int r = 0; r < list.Count; r++)
            {
                IRow row = sheet1.CreateRow(r + 1);
                for (int j = 0; j < properties.Count; j++)
                {
                    ICell cell = row.CreateCell(j);
                    var value = list[r].GetType().GetProperty(properties[j].Name)?.GetValue(list[r]);
                    if (value is DateTime)
                        value = ((DateTime)value).ToShortDateString();
                    cell.SetCellValue(value.ToString());
                }
            }

            using (var exportData = new MemoryStream())
            {
                workbook.Write(exportData);
                return exportData.GetBuffer();
            }
        }
        public static List<List<T>> Split<T>(List<T> collection, int size)
        {
            var chunks = new List<List<T>>();
            var chunkCount = collection.Count() / size;

            if (collection.Count % size > 0)
                chunkCount++;

            for (var i = 0; i < chunkCount; i++)
                chunks.Add(collection.Skip(i * size).Take(size).ToList());

            return chunks;
        }
    }
}

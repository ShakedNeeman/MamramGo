using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShowApp
{
    internal class MyExcel
    {

        static public void ExportToExcelSheet<T>(Dictionary<string, T> data, string sheetName, ExcelPackage excel, string[] headers)
        {
            var worksheet = excel.Workbook.Worksheets.Add(sheetName);

            // Add headers to the worksheet
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
            }

            // Initialize the row counter
            int row = 2;

            // Loop through each entry in the dictionary
            foreach (var entry in data)
            {
                // Add the key to the first column of the current row
                worksheet.Cells[row, 1].Value = entry.Key;

                // Check the type of the value
                if (entry.Value is Tuple<string, string, string, string> tupleValue)
                {
                    // Manually unpack the tuple and add the items to the worksheet
                    worksheet.Cells[row, 2].Value = tupleValue.Item1;
                    worksheet.Cells[row, 3].Value = tupleValue.Item2;
                    worksheet.Cells[row, 4].Value = tupleValue.Item3;
                    worksheet.Cells[row, 5].Value = tupleValue.Item4;
                }
                else if (entry.Value is Tuple<EventLogEntryType, DateTime, string, int> TValue)
                {
                    // Manually unpack the tuple and add the items to the worksheet
                    worksheet.Cells[row, 2].Value = TValue.Item1;
                    worksheet.Cells[row, 3].Value = TValue.Item2;
                    worksheet.Cells[row, 4].Value = TValue.Item3;
                    worksheet.Cells[row, 5].Value = TValue.Item4;
                }
                else if (entry.Value is ValueTuple<string, string> valueTuple)
                {
                    // Manually unpack the value tuple and add the items to the worksheet
                    worksheet.Cells[row, 2].Value = valueTuple.Item1;
                    worksheet.Cells[row, 3].Value = valueTuple.Item2;
                }
                else if(entry.Value is Tuple<string, DateTime, long> appValue )
                {
                    // Manually unpack the tuple and add the items to the worksheet
                    worksheet.Cells[row, 2].Value = appValue.Item1;
                    worksheet.Cells[row, 3].Value = appValue.Item2;
                    worksheet.Cells[row, 4].Value = appValue.Item3;

                }
                else
                {
                    // Convert the value to a string and add it to the second column of the current row
                    worksheet.Cells[row, 2].Value = entry.Value.ToString();
                }

                // Increment the row counter
                row++;
            }
        }

        public static void  ExportDataToExcel(Dictionary<string, object> appDetails, string sheetName, ExcelPackage excel, string[] headers)
        {
            ExportToExcelSheet(appDetails, sheetName, excel, headers);
        }

        public static void SaveExcelFile(ExcelPackage excel, string fileName)
        {
            FileInfo excelFile = new FileInfo(fileName);
            try
            {
                excel.SaveAs(excelFile);
                Console.WriteLine($"Excel file has been saved at: {excelFile.FullName}");
            }
            catch (Exception e)
            {
                HandleExceptions(e);
            }
        }

        public static void HandleExceptions(Exception e)
        {
            switch (e)
            {
                case IOException ioEx:
                    Console.WriteLine($"IO Error: {ioEx.Message}");
                    break;
                case UnauthorizedAccessException unAuthEx:
                    Console.WriteLine($"Access Error: You don't have permission to write to this location. {unAuthEx.Message}");
                    break;
                default:
                    Console.WriteLine($"An unknown error occurred: {e.Message}");
                    break;
            }
        }

        public static void CreateSummarySheet(ExcelPackage package, Dictionary<string, int> appCount, string sheetName)
        {
            var summaryWorksheet = package.Workbook.Worksheets.Add(sheetName);
            summaryWorksheet.Cells[1, 1].Value = "App Name";
            summaryWorksheet.Cells[1, 2].Value = "Count";

            int summaryRow = 2;
            foreach (var kvp in appCount)
            {
                summaryWorksheet.Cells[summaryRow, 1].Value = kvp.Key;
                summaryWorksheet.Cells[summaryRow, 2].Value = kvp.Value;
                summaryRow++;
            }

            package.Save();
            Console.WriteLine($"{sheetName} has been added to the Excel file.");
        }

    }
}

using System;
using System.IO;
using System.Linq;
using System.Text;
using GemBox.Spreadsheet;



namespace DataConverter
{
    public class Spreadsheet
    {

        public string path;
        public string format;

        public Spreadsheet(string pathFile, string format)
        {
            path = pathFile;

            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        }


        public ExcelFile LoadExcelFile()
        {
            return ExcelFile.Load(path);
        }

        public void Results()
        {

            // Iterate through all worksheets in a workbook.
            foreach (ExcelWorksheet worksheet in LoadExcelFile().Worksheets)
            {
                // Display sheet's name.
                Console.WriteLine("{1} {0} {1}\n", worksheet.Name, new string('#', 30));

                // Iterate through all rows in a worksheet.
                foreach (ExcelRow row in worksheet.Rows)
                {
                    // Iterate through all allocated cells in a row.
                    foreach (ExcelCell cell in row.AllocatedCells)
                    {
                        // Read cell's data.
                        string value = cell.Value?.ToString() ?? "EMPTY";

                        // Display cell's value and type.
                        value = value.Length > 15 ? value.Remove(15) + "…" : value;
                        Console.Write($"{value} [{cell.ValueType}]".PadRight(30));
                    }

                    Console.WriteLine();
                }


            }
        }
    }
}

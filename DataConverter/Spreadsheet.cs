﻿using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace DataConverter
{
    public class Spreadsheet
    {
        enum Format
        {
            xlsx,
            xls
        }

        public string filePath;
        public string format;

        public FileInfo existingFile;

    
        //constructor
        public Spreadsheet(string pathFile, string inputformat)
        {
            try
            {
                filePath = pathFile;
                format = inputformat;

                CheckInputFormat();

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                existingFile = new FileInfo(filePath);

            }
            catch (Exception ex)
            {
                throw ex;

            }

        }

        //check format if valid
        public void CheckInputFormat()
        {

            if(!Enum.IsDefined(typeof(Format), format))
            {
                throw new Exception("Formato non gestito");
            }          

        }


        //read data and output in List<string[]>
        public List<string[]> ReadData()
        {

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //Get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                List<string[]> listRow = new List<string[]>();
                List<string> header = new List<string>();
                
                for (int row = 1; row <= rowCount; row++)
                {
                    
                    //row array of strings
                    string[] rowString= new string[colCount];


                    for(int col = 1; col<= colCount; col++)
                    {                       
                        //cell value
                        rowString[col-1] = worksheet.Cells[row, col].Value.ToString().Trim();
                        //Console.WriteLine(" Row:" + row + " column:" + col + " Value:" +    worksheet.Cells[row, col].Value.ToString().Trim());
                    }

                    //add row to list
                    listRow.Add(rowString);
                }

            return listRow;
            } // the using statement automatically calls Dispose() which closes the package.
        }




        
    }
}

using System;
using System.Text.Json;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Json.More;

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

        public List<string[]> rowsList;
    
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
               throw new Exception("Formato non gestito");  
        }


        //read data and output in List<string[]>
        public void ReadData()
        {

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //Get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count

                rowsList = new List<string[]>();
                
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
                    rowsList.Add(rowString);
                }

            }
        }

        public string ConvertToJson()
        {
         
            //list of dictionary [{}]
            List<Dictionary<string, JsonElement>> listDict = new List<Dictionary<string, JsonElement>>();
            for (int row = 0; row < rowsList.Count; row++)
            {
                //from second row (first row is header)
                if (row == 0)
                {
                    continue;
                }

                Dictionary<string, JsonElement> dict = new Dictionary<string, JsonElement>();


                for (int col = 0; col < rowsList[row].Length; col++)
                {
                    //check and parse datatype

                    dict.Add(rowsList[0][col], rowsList[row][col].AsJsonElement());
                }

                listDict.Add(dict);
            }

            var json = JsonSerializer.Serialize(listDict);

            return json;
        }

        public static string GenerateMap(List<string[]> list)
        {
            //starting from one row, analyze it and then compose a json datatype map
            string map = "";

            //take first row
            string[] firstRow = list[1];

            Dictionary<string, string> dict= new Dictionary<string, string>();
            List<Dictionary<string, string>> listDict = new List<Dictionary<string, string>>();

            //foreach cell value try to parse to int or float
            for (int i = 0; i < firstRow.Length; i++)
            {
                string datatype;
                //float
                if (Double.TryParse(firstRow[i], out double result) || Int32.TryParse(firstRow[i],out int integer))
                    datatype = "number";
                //int
                else
                    //Int32.TryParse(firstRow[i], out int result);
                    datatype = "string";

                dict.Add("Type",datatype);

            }

            return map;
             
        }


        
    }
}

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

        enum Datatypes
        {
            Int,
            Double,
            Decimal,
            String,
            Float
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

        //public string ConvertToJson()
        //{

        //    //generate map
        //    List<Dictionary<string, JsonValueKind>> mapDict = GenerateMap(rowsList);

        //    //list of dictionary [{}]
        //    List<Dictionary<string, JsonElement>> listDict = new List<Dictionary<string, JsonElement>>();
        //    for (int row = 0; row < rowsList.Count; row++)
        //    {
        //        //from second row (first row is header)
        //        if (row == 0)
        //        {
        //            continue;
        //        }

        //        Dictionary<string, JsonElement> dict = new Dictionary<string, JsonElement>();

        //        for (int col = 0; col < rowsList[row].Length; col++)
        //        {
        //            //dynamic type cell value, don't know if string or int yet
        //            //dynamic cell = rowsList[row][col];

        //            switch (mapDict[col]["Type"])
        //            {
        //                case JsonValueKind.Number:
        //                    dict.Add(rowsList[0][col], GetInt(rowsList[row][col]).AsJsonElement());
        //                    break;

        //                default:
        //                    //add to dict
        //                    dict.Add(rowsList[0][col], rowsList[row][col].AsJsonElement());
        //                    break;
        //            }

        //        }

        //        listDict.Add(dict);
        //    }

        //    var json = JsonSerializer.Serialize(listDict);

        //    return json;
        //}

        public string ConvertToJson()
        {

            //generate map
            List<Dictionary<string, Datatypes>> mapDict = GenerateMap(rowsList);

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
                    //dynamic type cell value, don't know if string or int yet
                    JsonElement value = ConvertNumber(rowsList[row][col], mapDict[col]["Type"]);
                    dict.Add(rowsList[0][col], value);
                }

                listDict.Add(dict);
            }

            var json = JsonSerializer.Serialize(listDict);

            return json;
        }

        private static JsonElement ConvertNumber(string value, Datatypes type)
        {
            JsonElement result = 1.AsJsonElement();

            switch (type)
            {
                case Datatypes.Int:
                     result = Convert.ToInt32(value).AsJsonElement();
                    break;
                case Datatypes.Double:
                    result = Convert.ToDouble(value).AsJsonElement();
                    break;
                case Datatypes.Float:
                    result = Convert.ToDecimal(value).AsJsonElement();
                    break;
                case Datatypes.String:
                    result = value.AsJsonElement();
                    break;
            }   
            return result;
        }


        //Generate conversion map
        //private static List<Dictionary<string, JsonValueKind>> GenerateMap(List<string[]> list)
        //{
        //    //starting from one row, analyze it and then compose a json datatype map

        //    //take first row
        //    string[] firstRow = list[1];

        //    List<Dictionary<string, JsonValueKind>> mapDict = new List<Dictionary<string, JsonValueKind>>();

        //    //foreach cell value try to parse to int or float
        //    for (int i = 0; i < firstRow.Length; i++)
        //    {
        //        Dictionary<string, JsonValueKind> dict= new Dictionary<string, JsonValueKind>();
        //        JsonValueKind datatype;
        //        //float
        //        if (Double.TryParse(firstRow[i], out double result) || Int32.TryParse(firstRow[i],out int integer))
        //            datatype = JsonValueKind.Number;
        //            //int
        //        else
        //            //Int32.TryParse(firstRow[i], out int result);
        //            datatype = JsonValueKind.String;

        //        dict.Add("Type",datatype);

        //        mapDict.Add(dict);
        //    }
        //    return mapDict;
        //}

        //Generate conversion map
        private static List<Dictionary<string, Datatypes>> GenerateMap(List<string[]> list)
        {
            //starting from one row, analyze it and then compose a json datatype map

            //take first row
            string[] firstRow = list[1];

            List<Dictionary<string, Datatypes>> mapDict = new List<Dictionary<string, Datatypes>>();

            //foreach cell value try to parse to int or float
            for (int i = 0; i < firstRow.Length; i++)
            {
                Dictionary<string, Datatypes> dict = new Dictionary<string, Datatypes>();
                Datatypes datatype;
                //float

                if (Int32.TryParse(firstRow[i], out int integer))
                    datatype = Datatypes.Int;
                else if(Double.TryParse(firstRow[i], out double result))
                    datatype = Datatypes.Double;
                else if (float.TryParse(firstRow[i], out float floatNumber))
                    datatype = Datatypes.Float;
                else
                    datatype = Datatypes.String;

                dict.Add("Type", datatype);

                mapDict.Add(dict);
            }
            return mapDict;
        }

    }
}

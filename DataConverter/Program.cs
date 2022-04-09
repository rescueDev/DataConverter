using System.Collections.Generic;
using System.Text.Json;
using Json.Patch;
using Json.Pointer;
using OfficeOpenXml;
using Json.More;
using System;

namespace DataConverter
{
    class Program
    {
        static void Main(string[] args)
        {


            Spreadsheet spreadsheet = new Spreadsheet("/Users/salvatoreborgia/Downloads/test.xlsx", "xlsx");

            List<string[]> listRow = spreadsheet.ReadData();

            //list of dictionary [{}]
            List<Dictionary<string, JsonElement>> listDict = new List<Dictionary<string, JsonElement>>();
            
            for (int row = 0; row < listRow.Count; row++)
            {
                //from second row (first row is header)
                if (row == 0) continue;
                
                Dictionary<string, JsonElement> dict = new Dictionary<string, JsonElement>();
            

                for (int col = 0; col < listRow[row].Length; col++)
                {
                    //check and parse datatype

                    dict.Add(listRow[0][col],listRow[row][col].AsJsonElement());
                }

                listDict.Add(dict);           
            }

            //conversione json data
            

            var json = JsonSerializer.Serialize(listDict);
            
                        
            //JsonPointer pointer = JsonPointer.Parse("/1/Nome");
            //string json = "[{\"Nome\": \"Salvatore\",\"Age\": 29}, {\"Nome\": \"Agostino\",\"Age\": 25}]";
            //JsonPatch patch = new JsonPatch(PatchOperation.Replace(pointer, "Pippo"));
            //JsonDocument result = patch.Apply(JsonDocument.Parse(json));
            //JsonElement jsonfinale = JsonSchema.ToJsonElement(result);
            //Console.WriteLine(json);

        }

    }
}

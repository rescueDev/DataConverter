using System.Collections.Generic;
using System.Text.Json;
using Json.Patch;
using Json.Pointer;
using OfficeOpenXml;
using Json.More;

namespace DataConverter
{
    class Program
    {
        static void Main(string[] args)
        {


            Spreadsheet spreadsheet = new Spreadsheet("/Users/salvatoreborgia/Downloads/test.xlsx", "xlsx");

            List<string[]> listRow = spreadsheet.ReadData();

            //list of dictionary [{}]
            List<Dictionary<string, string>> listDict = new List<Dictionary<string, string>>();

            for (int row = 0; row < listRow.Count; row++)
            {
                //from second row (first row is header)
                if (row == 0) continue;
                
                Dictionary<string, string> dict = new Dictionary<string, string>();

                for (int col = 0; col < listRow[row].Length; col++)
                {
                    
                    dict.Add(listRow[0][col],listRow[row][col]);
                }

                listDict.Add(dict);
                
            }
      
            var json = JsonSerializer.Serialize(listDict);

            
            //JsonPointer pointer = JsonPointer.Parse("/1/Nome");
            //string json = "[{\"Nome\": \"Salvatore\",\"Age\": 29}, {\"Nome\": \"Agostino\",\"Age\": 25}]";
            //JsonPatch patch = new JsonPatch(PatchOperation.Replace(pointer, "Pippo"));
            //JsonDocument result = patch.Apply(JsonDocument.Parse(json));
            //JsonElement jsonfinale = JsonSchema.ToJsonElement(result);
            //Console.WriteLine(jsonfinale);

        }

    }
}

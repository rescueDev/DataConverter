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


            var spreadsheet = new Spreadsheet("/Users/salvatoreborgia/Downloads/test.xlsx", "xlsx");
            spreadsheet.ReadData();
            string json = spreadsheet.ConvertToJson();

            //JsonPointer pointer = JsonPointer.Parse("/1/Nome");
            //string json = "[{\"Nome\": \"Salvatore\",\"Age\": 29}, {\"Nome\": \"Agostino\",\"Age\": 25}]";
            //JsonPatch patch = new JsonPatch(PatchOperation.Replace(pointer, "Pippo"));
            //JsonDocument result = patch.Apply(JsonDocument.Parse(json));
            //JsonElement jsonfinale = JsonSchema.ToJsonElement(result);
            Console.WriteLine(json);

        }

    }
}

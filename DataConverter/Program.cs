using System;
using System.Text;
using System.Text.Json;
using Json.Patch;
using Json.Pointer;
using Microsoft.Office.Interop.Excel;

namespace DataConverter
{
    class Program
    {
        static void Main(string[] args)
        {

            Spreadsheet excel = new Spreadsheet("/Users/salvatoreborgia/Downloads/test.xls", "xls");

            excel.Results();


            //JsonPointer pointer = JsonPointer.Parse("/1/Nome");
            //string json = "[{\"Nome\": \"Salvatore\",\"Age\": 29}, {\"Nome\": \"Agostino\",\"Age\": 25}]";
            //JsonPatch patch = new JsonPatch(PatchOperation.Replace(pointer, "Pippo"));
            //JsonDocument result = patch.Apply(JsonDocument.Parse(json));
            //JsonElement jsonfinale = JsonSchema.ToJsonElement(result);
            //Console.WriteLine(jsonfinale);

        }

     }
}

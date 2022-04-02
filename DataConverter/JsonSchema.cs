using System.Text.Json;
using Json.Patch;

namespace DataConverter
{ 
    public class JsonSchema
    {
        public JsonSchema()
        {



        }

        public static JsonElement ToJsonElement(JsonDocument document)
        {
            return JsonSerializer.Deserialize<JsonElement>(document);
        }

        //public JsonPatch ApplyPatch()
        //{
        //    JsonPatch patch = "";
        //    return patch;
        //}

        public string GeneratePatchOperation()
        {
            string patch = "";

            return patch;
                       
        }
    }
}

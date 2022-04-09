using System.Text.Json;
using Json.Patch;
using Json.More;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DataConverter
{
    public class JsonSchema
    {

        private readonly JsonElement _raw;

        public Dictionary<string, JsonSchema>? Object { get; }
        public List<JsonSchema>? Array { get; }

        //constructor JsonS
        public JsonSchema(JsonElement raw)
        {

            switch (raw.ValueKind)
            {
                case JsonValueKind.Object:
                    Object = raw.EnumerateObject()
                        .ToDictionary(kvp => kvp.Name, kvp => new JsonSchema(kvp.Value));
                    break;
                case JsonValueKind.Array:
                    Array = raw.EnumerateArray()
                        .Select(v => new JsonSchema(v))
                        .ToList();
                    break;
                case JsonValueKind.Undefined:
                case JsonValueKind.String:
                case JsonValueKind.Number:
                case JsonValueKind.True:
                case JsonValueKind.False:
                case JsonValueKind.Null:
                    _raw = raw.Clone();
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
      
    

        //build json element from string json
        public static JsonElement Build(string json)
        {
            JsonDocument doc = JsonDocument.Parse(json);
            return doc.RootElement.Clone();
        }


        //convert to element
        public JsonElement ToElement()
        {
            using var doc = JsonDocument.Parse(ToString());
            return doc.RootElement.Clone();
        }


        //output to string 
        public override string ToString()
        {
            if (Object != null)
                return $"{{{string.Join(",", Object.Select(p => $"{JsonSerializer.Serialize(p.Key)}:{p.Value}"))}}}";

            if (Array != null)
                return $"[{string.Join(",", Array)}]";

            return _raw.ValueKind switch
            {
                JsonValueKind.String => _raw.GetRawText(),
                JsonValueKind.Number => _raw.ToString(),
                JsonValueKind.True => "true",
                JsonValueKind.False => "false",
                JsonValueKind.Null => "null",
                _ => throw new ArgumentOutOfRangeException()
            };
        }

        //shortcut to json element deserialization
        public static JsonElement ToJsonElement(JsonDocument document)
        {
            return JsonSerializer.Deserialize<JsonElement>(document);
        }

        //generate patch operation
        public string GeneratePatchOperation()
        {
            string patch = "";

            return patch;

        }

    }

}


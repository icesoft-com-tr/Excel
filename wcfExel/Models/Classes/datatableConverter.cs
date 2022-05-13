using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace wcfExel.Models.Classes
{
    public class datatableConverter : JsonConverter
    {
        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(DataTable);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            JArray array = JArray.Load(reader);
            var dataTypes = DetermineColumnDataTypes(array);
            var table = BuildDataTable(array, dataTypes);
            return table;
        }

        private DataTable BuildDataTable(JArray array, Dictionary<string, Type> dataTypes)
        {
            DataTable table = new DataTable();
            foreach (var kvp in dataTypes)
            {
                table.Columns.Add(kvp.Key, kvp.Value);
            }

            foreach (JObject item in array.Children<JObject>())
            {
                DataRow row = table.NewRow();
                foreach (JProperty prop in item.Properties())
                {
                    if (prop.Value.Type != JTokenType.Null)
                    {
                        Type dataType = dataTypes[prop.Name];
                        row[prop.Name] = prop.Value.ToObject(dataType);
                    }
                }
                table.Rows.Add(row);
            }
            return table;
        }

        private Dictionary<string, Type> DetermineColumnDataTypes(JArray array)
        {
            var dataTypes = new Dictionary<string, Type>();
            foreach (JObject item in array.Children<JObject>())
            {
                foreach (JProperty prop in item.Properties())
                {
                    Type currentType = GetDataType(prop.Value.Type);
                    if (currentType != null)
                    {
                        Type previousType;
                        if (!dataTypes.TryGetValue(prop.Name, out previousType) ||
                            (previousType == typeof(long) && currentType == typeof(decimal)))
                        {
                            dataTypes[prop.Name] = currentType;
                        }
                        else if (previousType != currentType)
                        {
                            dataTypes[prop.Name] = typeof(string);
                        }
                    }
                }
            }
            return dataTypes;
        }

        private Type GetDataType(JTokenType tokenType)
        {
            switch (tokenType)
            {
                case JTokenType.Null:
                    return null;
                case JTokenType.String:
                    return typeof(string);
                case JTokenType.Integer:
                    return typeof(int);
                case JTokenType.Float:
                    return typeof(decimal);
                case JTokenType.Boolean:
                    return typeof(bool);
                case JTokenType.Date:
                    return typeof(DateTime);
                case JTokenType.TimeSpan:
                    return typeof(TimeSpan);
                case JTokenType.Guid:
                    return typeof(Guid);
                case JTokenType.Bytes:
                    return typeof(byte[]);
                case JTokenType.Array:
                case JTokenType.Object:
                    throw new JsonException("This converter does not support complex types");
                default:
                    return typeof(string);
            }
        }

        public override bool CanWrite
        {
            get { return false; }
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }
    }
}
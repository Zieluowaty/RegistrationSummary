using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Models.Interfaces;

namespace RegistrationSummary.Maui.Converters;

public class ProductConverter : JsonConverter
{
    public override bool CanConvert(Type objectType)
    {
        return objectType == typeof(IProduct);
    }

    public override object ReadJson(JsonReader reader, Type objectType, object? existingValue, JsonSerializer serializer)
    {
        JObject jsonObject = JObject.Load(reader);
        var typeName = jsonObject["Type"].ToString();

        IProduct product;

        if (typeName == "Course")
        {
            product = new Course();
        }
        else
        {
            throw new NotSupportedException($"Unsupported product type: {typeName}");
        }

        serializer.Populate(jsonObject.CreateReader(), product);
        return product;
    }

    public override void WriteJson(JsonWriter writer, object? value, JsonSerializer serializer)
    {
        throw new NotImplementedException();
    }
}

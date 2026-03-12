using System.Text.Json;
using System.Text.Json.Serialization;

namespace PaperFormat.Infrastructure.Word;

internal static class JsonOptionsFactory
{
    public static JsonSerializerOptions Create()
    {
        return new JsonSerializerOptions(JsonSerializerDefaults.Web)
        {
            WriteIndented = true,
            Converters =
            {
                new JsonStringEnumConverter()
            }
        };
    }
}

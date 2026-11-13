using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using Tomlyn;
using Tomlyn.Model;

namespace OfficeIMO.Data.Conversions.Benchmarks;

/// <summary>
/// The TOML pipelines people build from Tomlyn and System.Text.Json: Tomlyn's model walked into <see cref="Utf8JsonWriter"/>,
/// and a <see cref="JsonDocument"/> mapped into Tomlyn's model for <see cref="TomlSerializer"/>. They use the same mapping
/// as OfficeIMO (date-times as RFC 3339 strings, nulls omitted, a root array wrapped in <c>root</c>).
/// </summary>
internal static class TomlynPipelines {
    private static readonly JsonWriterOptions Indented = new() { Indented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    public static byte[] ToJson(string toml) {
        TomlTable table = TomlSerializer.Deserialize<TomlTable>(toml)!;
        using var output = new MemoryStream();
        using (var writer = new Utf8JsonWriter(output, Indented)) Write(writer, table);
        return output.ToArray();
    }

    private static void Write(Utf8JsonWriter writer, object? value) {
        switch (value) {
            case TomlTable table:
                writer.WriteStartObject();
                foreach (KeyValuePair<string, object> entry in table) {
                    writer.WritePropertyName(entry.Key);
                    Write(writer, entry.Value);
                }

                writer.WriteEndObject();
                break;
            case TomlTableArray tables:
                writer.WriteStartArray();
                foreach (TomlTable item in tables) Write(writer, item);
                writer.WriteEndArray();
                break;
            case TomlArray array:
                writer.WriteStartArray();
                foreach (object? item in array) Write(writer, item);
                writer.WriteEndArray();
                break;
            case string text:
                writer.WriteStringValue(text);
                break;
            case long integer:
                writer.WriteNumberValue(integer);
                break;
            case double number:
                writer.WriteNumberValue(number);
                break;
            case bool flag:
                writer.WriteBooleanValue(flag);
                break;
            default:
                writer.WriteStringValue(Convert.ToString(value, CultureInfo.InvariantCulture));
                break;
        }
    }

    public static string FromJson(byte[] json) {
        using JsonDocument document = JsonDocument.Parse(json);
        object? root = Map(document.RootElement);
        TomlTable table = root as TomlTable ?? new TomlTable { ["root"] = root! };
        return TomlSerializer.Serialize(table);
    }

    private static object? Map(JsonElement element) {
        switch (element.ValueKind) {
            case JsonValueKind.Object:
                var table = new TomlTable();
                foreach (JsonProperty property in element.EnumerateObject()) {
                    object? value = Map(property.Value);
                    if (value != null) table[property.Name] = value;
                }

                return table;
            case JsonValueKind.Array:
                bool tables = element.GetArrayLength() > 0;
                foreach (JsonElement item in element.EnumerateArray()) tables &= item.ValueKind == JsonValueKind.Object;
                if (tables) {
                    var array = new TomlTableArray();
                    foreach (JsonElement item in element.EnumerateArray()) array.Add((TomlTable)Map(item)!);
                    return array;
                }

                var values = new TomlArray();
                foreach (JsonElement item in element.EnumerateArray()) {
                    object? value = Map(item);
                    if (value != null) values.Add(value);
                }

                return values;
            case JsonValueKind.String:
                return element.GetString();
            case JsonValueKind.Number:
                return element.TryGetInt64(out long integer) ? integer : element.GetDouble();
            case JsonValueKind.True:
                return true;
            case JsonValueKind.False:
                return false;
            default:
                return null;
        }
    }
}

using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Pdf;

namespace OfficeIMO;

// Keep persistence in the batch owner; PDF geometry has no JSON runtime requirement.
internal sealed class OfficeBatchPageSizeConverter : JsonConverter<PdfPageSize> {
    public override PdfPageSize Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) {
        using var document = JsonDocument.ParseValue(ref reader);
        var root = document.RootElement;
        RequireProperties(root, "width", "height");
        return new PdfPageSize(root.GetProperty("width").GetDouble(), root.GetProperty("height").GetDouble());
    }

    public override void Write(Utf8JsonWriter writer, PdfPageSize value, JsonSerializerOptions options) {
        writer.WriteStartObject();
        writer.WriteNumber("width", value.Width);
        writer.WriteNumber("height", value.Height);
        writer.WriteEndObject();
    }

    internal static void RequireProperties(JsonElement root, params string[] names) {
        if (root.ValueKind != JsonValueKind.Object || root.EnumerateObject().Count() != names.Length ||
            names.Any(name => !root.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.Number))
            throw new JsonException("Saved page geometry requires exactly its named numeric dimensions.");
    }
}

internal sealed class OfficeBatchPageMarginsConverter : JsonConverter<PdfPageMargins> {
    public override PdfPageMargins Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) {
        using var document = JsonDocument.ParseValue(ref reader);
        var root = document.RootElement;
        OfficeBatchPageSizeConverter.RequireProperties(root, "left", "top", "right", "bottom");
        return new PdfPageMargins(root.GetProperty("left").GetDouble(), root.GetProperty("top").GetDouble(),
            root.GetProperty("right").GetDouble(), root.GetProperty("bottom").GetDouble());
    }

    public override void Write(Utf8JsonWriter writer, PdfPageMargins value, JsonSerializerOptions options) {
        writer.WriteStartObject();
        writer.WriteNumber("left", value.Left);
        writer.WriteNumber("top", value.Top);
        writer.WriteNumber("right", value.Right);
        writer.WriteNumber("bottom", value.Bottom);
        writer.WriteEndObject();
    }
}

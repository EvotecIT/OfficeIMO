using System.Text.Json;

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookScriptImporter {
    private static StringReader OpenScript(ReadOnlyMemory<byte> input) => new(NotebookImportContext.DecodeUtf8(input));

    private static JsonDocumentOptions JsonOptions(NotebookImportContext context) => new() {
        MaxDepth = context.Limits.MaxJsonDepth
    };

    private static JsonElement ReadObjectProperty(JsonElement element, string name) =>
        element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out JsonElement property) ? property : default;

    private static string? ReadStringProperty(JsonElement element, string name) =>
        ReadObjectProperty(element, name) is { ValueKind: JsonValueKind.String } property ? property.GetString() : null;

    private static void SetLanguageInfo(NotebookDocument document, string language) {
        if (document.Metadata.ContainsKey("language_info")) return;
        using var json = new MemoryStream();
        using (var writer = new Utf8JsonWriter(json)) {
            writer.WriteStartObject();
            writer.WriteString("name", language);
            writer.WriteEndObject();
        }
        using var value = JsonDocument.Parse(json.GetBuffer().AsMemory(0, (int)json.Length));
        document.SetMetadata("language_info", value.RootElement);
    }
}

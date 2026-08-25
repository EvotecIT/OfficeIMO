using System.Text.Json;

namespace OfficeIMO.Adf.Benchmarks;

internal sealed record AdfBenchmarkScale(string Name, int Paragraphs);

internal static class AdfBenchmarkCorpus {
    internal static readonly IReadOnlyList<AdfBenchmarkScale> Scales = [
        new("Small", 25),
        new("Normal", 500)
    ];

    internal static AdfBenchmarkScale Get(string name) =>
        Scales.FirstOrDefault(scale => string.Equals(scale.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException($"Unknown scale '{name}'.", nameof(name));

    internal static string Create(AdfBenchmarkScale scale) {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream)) {
            writer.WriteStartObject();
            writer.WriteNumber("version", 1);
            writer.WriteString("type", "doc");
            writer.WritePropertyName("content");
            writer.WriteStartArray();
            for (int index = 0; index < scale.Paragraphs; index++) {
                WriteParagraph(writer, index);
                if (index % 10 == 0) WriteList(writer, index);
                if (index % 25 == 0) WriteTable(writer, index);
            }
            writer.WriteStartObject();
            writer.WriteString("type", "futureNode");
            writer.WritePropertyName("attrs");
            writer.WriteStartObject();
            writer.WriteNumber("revision", scale.Paragraphs);
            writer.WriteEndObject();
            writer.WriteString("futurePayload", "preserved");
            writer.WriteEndObject();
            writer.WriteEndArray();
            writer.WritePropertyName("sourceExtension");
            writer.WriteStartObject();
            writer.WriteString("owner", "OfficeIMO benchmark");
            writer.WriteNumber("paragraphs", scale.Paragraphs);
            writer.WriteEndObject();
            writer.WriteEndObject();
        }
        return System.Text.Encoding.UTF8.GetString(stream.GetBuffer(), 0, checked((int)stream.Length));
    }

    private static void WriteParagraph(Utf8JsonWriter writer, int index) {
        writer.WriteStartObject();
        writer.WriteString("type", index % 10 == 0 ? "heading" : "paragraph");
        if (index % 10 == 0) {
            writer.WritePropertyName("attrs");
            writer.WriteStartObject();
            writer.WriteNumber("level", 2);
            writer.WriteEndObject();
        }
        writer.WritePropertyName("content");
        writer.WriteStartArray();
        WriteText(writer, $"Paragraph {index:D6} alpha beta gamma ", "strong", null);
        WriteText(writer, "linked text", "link", $"https://example.test/{index}");
        WriteText(writer, " complete", null, null);
        writer.WriteEndArray();
        writer.WritePropertyName("vendorNode");
        writer.WriteStartObject();
        writer.WriteNumber("rank", index);
        writer.WriteBoolean("retained", true);
        writer.WriteEndObject();
        writer.WriteEndObject();
    }

    private static void WriteText(Utf8JsonWriter writer, string text, string? mark, string? href) {
        writer.WriteStartObject();
        writer.WriteString("type", "text");
        writer.WriteString("text", text);
        if (mark != null) {
            writer.WritePropertyName("marks");
            writer.WriteStartArray();
            writer.WriteStartObject();
            writer.WriteString("type", mark);
            if (href != null) {
                writer.WritePropertyName("attrs");
                writer.WriteStartObject();
                writer.WriteString("href", href);
                writer.WriteString("title", "benchmark");
                writer.WriteEndObject();
            }
            writer.WriteString("vendorMarkFlag", "preserved");
            writer.WriteEndObject();
            writer.WriteEndArray();
        }
        writer.WriteEndObject();
    }

    private static void WriteList(Utf8JsonWriter writer, int index) {
        writer.WriteStartObject();
        writer.WriteString("type", "bulletList");
        writer.WritePropertyName("content");
        writer.WriteStartArray();
        for (int item = 0; item < 3; item++) {
            writer.WriteStartObject();
            writer.WriteString("type", "listItem");
            writer.WritePropertyName("content");
            writer.WriteStartArray();
            writer.WriteStartObject();
            writer.WriteString("type", "paragraph");
            writer.WritePropertyName("content");
            writer.WriteStartArray();
            WriteText(writer, $"List {index:D6}/{item}", null, null);
            writer.WriteEndArray();
            writer.WriteEndObject();
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    private static void WriteTable(Utf8JsonWriter writer, int index) {
        writer.WriteStartObject();
        writer.WriteString("type", "table");
        writer.WritePropertyName("content");
        writer.WriteStartArray();
        for (int row = 0; row < 2; row++) {
            writer.WriteStartObject();
            writer.WriteString("type", "tableRow");
            writer.WritePropertyName("content");
            writer.WriteStartArray();
            for (int column = 0; column < 3; column++) {
                writer.WriteStartObject();
                writer.WriteString("type", row == 0 ? "tableHeader" : "tableCell");
                writer.WritePropertyName("content");
                writer.WriteStartArray();
                writer.WriteStartObject();
                writer.WriteString("type", "paragraph");
                writer.WritePropertyName("content");
                writer.WriteStartArray();
                WriteText(writer, $"Cell {index:D6}/{row}/{column}", null, null);
                writer.WriteEndArray();
                writer.WriteEndObject();
                writer.WriteEndArray();
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
        writer.WriteEndObject();
    }
}

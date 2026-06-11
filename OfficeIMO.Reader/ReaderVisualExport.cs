using System;
using System.IO;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Reader;

/// <summary>
/// Deterministic text exporters for <see cref="ReaderVisual"/> instances.
/// </summary>
public static class ReaderVisualExport {
    /// <summary>
    /// Serializes a reader visual as deterministic JSON.
    /// </summary>
    /// <param name="visual">Visual to serialize.</param>
    /// <param name="indented">When true, writes indented JSON for diagnostics and fixtures.</param>
    public static string ToJson(this ReaderVisual visual, bool indented = false) {
        if (visual == null) throw new ArgumentNullException(nameof(visual));

        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = indented })) {
            writer.WriteStartObject();
            ReaderJsonWriter.WriteNullableString(writer, "kind", visual.Kind);
            ReaderJsonWriter.WriteNullableString(writer, "language", visual.Language);
            ReaderJsonWriter.WriteNullableString(writer, "payloadHash", visual.PayloadHash);
            ReaderJsonWriter.WriteNullableString(writer, "sourceName", visual.SourceName);
            ReaderJsonWriter.WriteNullableString(writer, "mimeType", visual.MimeType);
            WriteNullableNumber(writer, "width", visual.Width);
            WriteNullableNumber(writer, "height", visual.Height);
            WriteNullableNumber(writer, "x", visual.X);
            WriteNullableNumber(writer, "y", visual.Y);
            WriteNullableNumber(writer, "placedWidth", visual.PlacedWidth);
            WriteNullableNumber(writer, "placedHeight", visual.PlacedHeight);
            writer.WriteNumber("placementCount", visual.PlacementCount);
            writer.WriteBoolean("hasGeometry", visual.HasGeometry);
            WriteNullableBoolean(writer, "isAxisAligned", visual.IsAxisAligned);
            ReaderJsonWriter.WriteNullableString(writer, "content", visual.Content);
            ReaderJsonWriter.WriteLocation(writer, visual.Location);
            writer.WriteEndObject();
        }

        return Encoding.UTF8.GetString(stream.ToArray());
    }

    private static void WriteNullableNumber(Utf8JsonWriter writer, string propertyName, double? value) {
        if (value.HasValue) {
            writer.WriteNumber(propertyName, value.Value);
        }
    }

    private static void WriteNullableBoolean(Utf8JsonWriter writer, string propertyName, bool? value) {
        if (value.HasValue) {
            writer.WriteBoolean(propertyName, value.Value);
        }
    }

    /// <summary>
    /// Returns a deterministic source-payload extension for a reader visual.
    /// </summary>
    /// <param name="visual">Visual whose kind and language should be inspected.</param>
    public static string GetPayloadExtension(ReaderVisual visual) {
        if (visual == null) throw new ArgumentNullException(nameof(visual));

        string language = NormalizeToken(visual.Language);
        if (language.Length == 0) {
            language = NormalizeToken(visual.Kind);
        }

        switch (language) {
            case "mermaid":
            case "mmd":
                return ".mmd";
            case "plantuml":
            case "puml":
                return ".puml";
            case "dot":
            case "graphviz":
                return ".dot";
            case "json":
            case "ix-chart":
            case "chart":
            case "network":
                return ".json";
            case "svg":
                return ".svg";
            default:
                return ".txt";
        }
    }

    private static string NormalizeToken(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        return value!.Trim().ToLowerInvariant();
    }

}

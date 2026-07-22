using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Reader;

/// <summary>Deterministic JSON serialization for structured extraction results.</summary>
public static class OfficeDocumentStructuredExtractionJson {
    /// <summary>Serializes a current structured extraction result.</summary>
    public static string Serialize(OfficeDocumentStructuredExtractionResult result, bool indented = false) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (!string.Equals(result.SchemaId, OfficeDocumentStructuredExtractionSchema.Id, StringComparison.Ordinal) ||
            result.SchemaVersion != OfficeDocumentStructuredExtractionSchema.Version) {
            throw new InvalidOperationException(
                $"Structured extraction schema '{result.SchemaId}' version {result.SchemaVersion} is not supported.");
        }
        var context = new ReaderJsonSerializerContext(CreateOptions(indented));
        return JsonSerializer.Serialize(result, context.OfficeDocumentStructuredExtractionResult);
    }

    /// <summary>Serializes a current structured extraction result.</summary>
    public static string ToJson(this OfficeDocumentStructuredExtractionResult result, bool indented = false) {
        return Serialize(result, indented);
    }

    private static JsonSerializerOptions CreateOptions(bool indented) {
        var options = new JsonSerializerOptions {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = indented
        };
        return options;
    }
}

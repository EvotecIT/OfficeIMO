using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Reader;

/// <summary>Deterministic JSON serialization for hierarchical chunking results.</summary>
public static class ReaderChunkHierarchyJson {
    /// <summary>Serializes a current hierarchy result.</summary>
    public static string Serialize(ReaderChunkHierarchyResult result, bool indented = false) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (!string.Equals(result.SchemaId, ReaderChunkHierarchySchema.Id, StringComparison.Ordinal) ||
            result.SchemaVersion != ReaderChunkHierarchySchema.Version) {
            throw new InvalidOperationException(
                $"Chunk hierarchy schema '{result.SchemaId}' version {result.SchemaVersion} is not supported.");
        }
        var options = new JsonSerializerOptions {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = indented
        };
        var context = new ReaderJsonSerializerContext(options);
        return JsonSerializer.Serialize(result, context.ReaderChunkHierarchyResult);
    }

    /// <summary>Serializes a current hierarchy result.</summary>
    public static string ToJson(this ReaderChunkHierarchyResult result, bool indented = false) =>
        Serialize(result, indented);
}

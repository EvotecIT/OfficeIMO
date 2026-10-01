using System.Text.Json;

namespace OfficeIMO.Workflows;

/// <summary>Source-generated JSON for bounded conversion batch summaries.</summary>
public static class OfficeConversionBatchSerializer {
    /// <summary>Serializes a bounded run summary for command-line consumers.</summary>
    public static string SerializeResult(OfficeConversionBatchResult result) =>
        JsonSerializer.Serialize(result, OfficeConversionBatchJsonContext.Default.OfficeConversionBatchResult);
}

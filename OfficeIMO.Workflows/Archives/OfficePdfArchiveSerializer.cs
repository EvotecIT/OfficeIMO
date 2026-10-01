using System.Text.Json;

namespace OfficeIMO.Workflows;

/// <summary>Strict, source-generated JSON for local archive requests and summaries.</summary>
public static class OfficePdfArchiveSerializer {
    /// <summary>Parses at most 64 KiB of UTF-8 JSON, rejecting unknown fields and missing requests.</summary>
    public static OfficePdfArchiveRequest ParseRequest(byte[] utf8Json) {
        ArgumentNullException.ThrowIfNull(utf8Json);
        if (utf8Json.Length > 64 * 1024) throw new InvalidDataException("Archive request exceeds 64 KiB.");
        return JsonSerializer.Deserialize(utf8Json, OfficePdfArchiveJsonContext.Default.OfficePdfArchiveRequest)
            ?? throw new InvalidDataException("An archive request is required.");
    }

    /// <summary>Serializes a request without persisting runtime credentials or accepting provider locations.</summary>
    public static string SerializeRequest(OfficePdfArchiveRequest request) =>
        JsonSerializer.Serialize(request, OfficePdfArchiveJsonContext.Default.OfficePdfArchiveRequest);

    /// <summary>Serializes a bounded run summary for command-line consumers.</summary>
    public static string SerializeResult(OfficePdfArchiveResult result) =>
        JsonSerializer.Serialize(result, OfficePdfArchiveJsonContext.Default.OfficePdfArchiveResult);
}

using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows;

/// <summary>Versioned text-review report transport with exact offsets and selected mutations.</summary>
public static class OfficeTextIntegrityReportSerializer {
    /// <summary>Exports immutable findings and selected indices; source changes are rejected by the review owner.</summary>
    public static string Serialize(OfficeTextIntegrityReview review, string currentText, string sourceName,
        IEnumerable<int>? selectedIndices = null) {
        ArgumentNullException.ThrowIfNull(review);
        int[] selected = (selectedIndices ?? Array.Empty<int>()).Distinct().OrderBy(index => index).ToArray();
        _ = review.RemoveSelected(currentText, selected);
        var document = new TextIntegrityDocument("officeimo.text-integrity.result.v1", sourceName, "Completed", "UTF-16 code units",
            review.TextSha256, review.InputSha256, review.EncodingName, review.HasByteOrderMark,
            selected.Length == 0 ? null : OfficeTextIntegrityReview.ComputeSha256(review.ExportSelected(currentText, selected)),
            review.Report.Findings.Select(item => new ProvenanceTextFindingDto(item.Kind.ToString(), item.Risk.ToString(),
                item.TextOffset, item.TextLength, item.CodePoint, item.UnicodeNotation, item.Location)).ToArray(), selected);
        return JsonSerializer.Serialize(document, TextIntegrityJsonContext.Default.TextIntegrityDocument);
    }
}
internal sealed record TextIntegrityDocument(string Schema, string SourceName, string Status, string OffsetUnit,
    string TextSha256, string? InputSha256, string Encoding, bool HasByteOrderMark, string? OutputSha256,
    IReadOnlyList<ProvenanceTextFindingDto> Findings, IReadOnlyList<int> SelectedFindingIndices);
[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(TextIntegrityDocument))]
internal sealed partial class TextIntegrityJsonContext : JsonSerializerContext;

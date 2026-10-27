using System.Text.Json.Serialization;
using OfficeIMO.Reader;

namespace OfficeIMO.AI;

// Named transport records keep the existing JSON shape and snapshot byte identity without runtime reflection.
internal sealed record OfficeAiSnapshotPayload(string sourceHash, OfficeAiEvidence[] evidence, int[] pages,
    string pageProvenance, bool hasSourceDiagnostics, IEnumerable<OfficeAiSnapshotImage> images);
internal sealed record OfficeAiSnapshotImage(string Id, int Page, string MediaType, int Width, int Height, string ContentHash);
internal sealed record OfficeAiReportPayload(string schema, OfficeAiReportSource source, IReadOnlyList<OfficeAiEvidence> evidence,
    IEnumerable<OfficeAiReportImage> images, OfficeAiResult result);
internal sealed record OfficeAiReportSource(string SourceHash, string SnapshotHash, int SourceByteLength,
    string PageProvenance, IReadOnlyList<int> Pages);
internal sealed record OfficeAiReportImage(string Id, int Page, string MediaType, int Width, int Height, int ByteLength, string ContentHash);
internal sealed record OfficeAiPlanningPayload(string schema, string sourceHash, string snapshotHash, string pageProvenance,
    string operation, string? instruction, string? conversationContext, OfficeAiResultLimits resultLimits,
    IEnumerable<OfficeAiPlanningField> fields, IEnumerable<OfficeAiPlanningEvidence> evidence, IEnumerable<OfficeAiPlanningImage> images);
internal sealed record OfficeAiResultLimits(int maxResultItems, int maxTableCells, int maxTableColumns);
internal sealed record OfficeAiPlanningField(string key, string name, string type, string? dateFormat);
internal sealed record OfficeAiPlanningEvidence(string id, string kind, string text, int? page,
    OfficeDocumentRecognitionEvidence? recognition, string? sourceBlockId, string? sourceAnchor, OfficeAiSourceLocation? sourceLocation);
internal sealed record OfficeAiPlanningImage(string id, int page, int width, int height);
internal sealed record OfficeAiSynthesisPayload(string schema, string operation, string? instruction, string? conversationContext,
    OfficeAiSynthesisCoverage coverage, int maxResultItems, IEnumerable<OfficeAiSynthesisDraft> drafts);
internal sealed record OfficeAiSynthesisCoverage(int omittedEvidenceCount, IReadOnlyList<int> emptyPages, bool sourceReaderDiagnostics);
internal sealed record OfficeAiSynthesisDraft(string id, string text, IEnumerable<OfficeAiSynthesisSource> sources);
internal sealed record OfficeAiSynthesisSource(string EvidenceId, int? Page, string? Quote, bool QuoteMatched,
    OfficeDocumentRecognitionEvidence? Recognition);

[JsonSourceGenerationOptions(GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(OfficeAiSnapshotPayload))]
[JsonSerializable(typeof(OfficeAiPlanningPayload))]
[JsonSerializable(typeof(OfficeAiSynthesisPayload))]
internal sealed partial class OfficeAiJsonContext : JsonSerializerContext;

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, WriteIndented = true,
    UseStringEnumConverter = true, GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(OfficeAiReportPayload))]
internal sealed partial class OfficeAiReportJsonContext : JsonSerializerContext;

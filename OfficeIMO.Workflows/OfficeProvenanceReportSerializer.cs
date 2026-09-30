using System.Text.Json;
using System.Text.Json.Serialization;
using System.Security.Cryptography;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows;

/// <summary>Canonical versioned provenance report transport for local and memory-only hosts.</summary>
public static class OfficeProvenanceReportSerializer {
    /// <summary>Serializes one immutable workflow result with string enums and explicit check states.</summary>
    public static string Serialize(OfficeProvenanceWorkflowResult result) =>
        JsonSerializer.Serialize(Create(result), ProvenanceReportJsonContext.Default.ProvenanceResultDto);

    /// <summary>Serializes a batch using the same item contract as single reports.</summary>
    public static string SerializeBatch(IReadOnlyList<OfficeProvenanceWorkflowResult> results) =>
        JsonSerializer.Serialize(new ProvenanceBatchDto("officeimo.provenance.batch.v2", results.Select(Create).ToArray()),
            ProvenanceReportJsonContext.Default.ProvenanceBatchDto);

    /// <summary>Creates a workflow result for already inspected memory-only bytes without reading a path.</summary>
    public static OfficeProvenanceWorkflowResult FromBuffer(string fileName, byte[] input, OfficeProvenanceReport inspection,
        OfficeProvenanceRemovalResult? removal = null) {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(inspection);
        return new OfficeProvenanceWorkflowResult(Guid.NewGuid().ToString("N"),
            removal == null ? OfficeProvenanceWorkflowOperation.Inspect : OfficeProvenanceWorkflowOperation.Remove,
            OfficeWorkflowStatus.Completed, OfficeWorkflowFailureKind.None,
            OfficeProvenanceWorkflowCatalog.FindByPath(fileName)?.OwnerPackage ?? "OfficeIMO.Core", null,
            input.LongLength, removal?.DataLength ?? 0, TimeSpan.Zero,
            removal == null ? "Structural inspection completed." : "Selected provenance carriers were processed in a separate copy.",
            Array.Empty<OfficeWorkflowDiagnostic>(), inspection: removal == null ? inspection : null,
            before: removal?.Before, after: removal?.After, changes: removal?.Changes,
            wasReserialized: removal?.WasReserialized ?? false,
            wereInvalidatedSignaturesRemoved: removal?.WereInvalidatedSignaturesRemoved ?? false,
            inputPath: fileName, inputSha256: Convert.ToHexString(SHA256.HashData(input)).ToLowerInvariant(),
            outputSha256: removal == null ? null : Convert.ToHexString(removal.ComputeDataSha256()).ToLowerInvariant());
    }

    /// <summary>Creates a transport document retaining evidence, diagnostic detail and check coverage.</summary>
    public static ProvenanceResultDto Create(OfficeProvenanceWorkflowResult result) => new(
        "officeimo.provenance.result.v2",
        result.InputPath, result.InputSha256, result.OutputSha256,
        OfficeProvenanceWorkflowCatalog.FindByPath(result.InputPath ?? "")?.Notes ?? "",
        new ProvenanceChecksDto(result.Succeeded ? "Completed" : "Failed",
            result.Assessment?.TextIntegrityStatus.ToString() ?? "NotRequested",
            result.Assessment?.VerificationStatus.ToString() ?? "NotRequested",
            result.Assessment?.ProviderSignalsStatus.ToString() ?? "NotRequested"),
        result.RequestId,
        result.Operation.ToString(),
        result.Status.ToString(),
        result.FailureKind.ToString(),
        result.OwnerPackage,
        result.OutputPath,
        result.InputBytes,
        result.OutputBytes,
        result.Duration.TotalMilliseconds,
        result.Summary,
        result.Inspection is null ? null : ToDto(result.Inspection),
        result.Assessment is null ? null : ToDto(result.Assessment),
        result.Before is null ? null : ToDto(result.Before),
        result.After is null ? null : ToDto(result.After),
        result.Changes.Select(change => new ProvenanceChangeDto(
            change.Carrier.ToString(), change.Location, change.RemovedBytes)).ToArray(),
        result.WasChanged,
        result.WasReserialized,
        result.WereInvalidatedSignaturesRemoved,
        result.Diagnostics.Select(diagnostic => new ProvenanceDiagnosticDto(
            diagnostic.Code,
            diagnostic.Message,
            diagnostic.Severity.ToString(),
            diagnostic.Stage,
            diagnostic.Details)).ToArray());

    private static ProvenanceAssessmentDto ToDto(OfficeProvenanceAssessmentReport report) => new(
        ToDto(report.Structural),
        report.TextIntegrityStatus.ToString(), report.VerificationStatus.ToString(), report.ProviderSignalsStatus.ToString(),
        report.Verification is null ? null : new ProvenanceVerificationDto(
            report.Verification.ProviderName,
            report.Verification.Status.ToString(),
            report.Verification.Findings,
            report.Verification.RawReport),
        report.TextIntegrity?.Findings.Select(finding => new ProvenanceTextFindingDto(
            finding.Kind.ToString(),
            finding.Risk.ToString(),
            finding.TextOffset,
            finding.TextLength,
            finding.CodePoint,
            finding.UnicodeNotation,
            finding.Location)).ToArray(),
        report.ProviderSignals.Select(signal => new ProvenanceSignalDto(
            signal.ProviderName,
            signal.SignalKind.ToString(),
            signal.Status.ToString(),
            signal.Findings)).ToArray());

    private static ProvenanceReportDto ToDto(OfficeProvenanceReport report) => new(
        report.Format.ToString(),
        report.Evidence.Select(evidence => new ProvenanceEvidenceDto(
            evidence.Carrier.ToString(),
            evidence.Location,
            evidence.IsStructurallyValid,
            evidence.PayloadLength,
            evidence.Value,
            evidence.DigitalSourceKind.ToString())).ToArray(),
        report.Diagnostics,
        report.HasC2paManifest,
        report.HasExternalC2paManifest,
        report.HasGenerativeAiDeclaration);
}

#pragma warning disable CS1591 // Transport fields are documented by their owning report contract.
public sealed record ProvenanceBatchDto(string Schema, IReadOnlyList<ProvenanceResultDto> Results);
public sealed record ProvenanceResultDto(
    string Schema,
    string? InputPath, string? InputSha256, string? OutputSha256, string CoverageNotes, ProvenanceChecksDto Checks,
    string RequestId,
    string Operation,
    string Status,
    string FailureKind,
    string OwnerPackage,
    string? OutputPath,
    long InputBytes,
    long OutputBytes,
    double DurationMilliseconds,
    string Summary,
    ProvenanceReportDto? Inspection,
    ProvenanceAssessmentDto? Assessment,
    ProvenanceReportDto? Before,
    ProvenanceReportDto? After,
    IReadOnlyList<ProvenanceChangeDto> Changes,
    bool WasChanged,
    bool WasReserialized,
    bool WereInvalidatedSignaturesRemoved,
    IReadOnlyList<ProvenanceDiagnosticDto> Diagnostics);
public sealed record ProvenanceReportDto(
    string Format,
    IReadOnlyList<ProvenanceEvidenceDto> Evidence,
    IReadOnlyList<string> Diagnostics,
    bool HasC2paManifest,
    bool HasExternalC2paManifest,
    bool HasGenerativeAiDeclaration);
public sealed record ProvenanceEvidenceDto(string Carrier, string Location, bool IsStructurallyValid, long PayloadLength, string? Value, string DigitalSourceKind);
public sealed record ProvenanceAssessmentDto(ProvenanceReportDto Structural, string TextIntegrityStatus, string VerificationStatus, string ProviderSignalsStatus, ProvenanceVerificationDto? Verification, IReadOnlyList<ProvenanceTextFindingDto>? TextIntegrity, IReadOnlyList<ProvenanceSignalDto> ProviderSignals);
public sealed record ProvenanceVerificationDto(string ProviderName, string Status, IReadOnlyList<string> Findings, string? RawReport);
public sealed record ProvenanceTextFindingDto(string Kind, string Risk, int TextOffset, int TextLength, int CodePoint, string UnicodeNotation, string Location);
public sealed record ProvenanceSignalDto(string ProviderName, string SignalKind, string Status, IReadOnlyList<string> Findings);
public sealed record ProvenanceChangeDto(string Carrier, string Location, long RemovedBytes);
public sealed record ProvenanceDiagnosticDto(string Code, string Message, string Severity, string? Stage, IReadOnlyDictionary<string, string> Details);

public sealed record ProvenanceChecksDto(string Structural, string TextIntegrity, string Verification, string ProviderSignals);
#pragma warning restore CS1591
[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull, GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(ProvenanceResultDto))]
[JsonSerializable(typeof(ProvenanceBatchDto))]
internal sealed partial class ProvenanceReportJsonContext : JsonSerializerContext;

using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Provenance;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Provenance;

internal static class ProvenanceOutput {
    internal static async Task WriteCapabilitiesAsync(TextWriter writer, ProvenanceOutputFormat format) {
        ProvenanceCapabilitiesDto dto = new(
            "officeimo.provenance.capabilities.v1",
            OfficeProvenanceWorkflowCatalog.All.Select(ToDto).ToArray());
        if (format == ProvenanceOutputFormat.Json) {
            await writer.WriteLineAsync(JsonSerializer.Serialize(dto, ProvenanceJsonContext.Default.ProvenanceCapabilitiesDto)).ConfigureAwait(false);
            return;
        }
        foreach (ProvenanceCapabilityDto capability in dto.Capabilities) {
            await writer.WriteLineAsync(
                capability.Id + " | " + capability.OwnerPackage + " | " +
                string.Join(',', capability.Extensions) + " | remove=" + capability.CanRemove.ToString().ToLowerInvariant()).ConfigureAwait(false);
        }
    }

    internal static async Task WriteResultAsync(
        TextWriter writer,
        OfficeProvenanceWorkflowResult result,
        ProvenanceOutputFormat format) {
        ProvenanceResultDto dto = ToDto(result);
        if (format == ProvenanceOutputFormat.Json) {
            await writer.WriteLineAsync(JsonSerializer.Serialize(dto, ProvenanceJsonContext.Default.ProvenanceResultDto)).ConfigureAwait(false);
            return;
        }
        await WriteTextResultAsync(writer, dto).ConfigureAwait(false);
    }

    internal static async Task WriteBatchAsync(
        TextWriter writer,
        IReadOnlyList<OfficeProvenanceWorkflowResult> results,
        ProvenanceOutputFormat format) {
        ProvenanceBatchDto dto = new(
            "officeimo.provenance.batch.v1",
            results.Select(ToDto).ToArray());
        if (format == ProvenanceOutputFormat.Json) {
            await writer.WriteLineAsync(JsonSerializer.Serialize(dto, ProvenanceJsonContext.Default.ProvenanceBatchDto)).ConfigureAwait(false);
            return;
        }
        foreach (ProvenanceResultDto result in dto.Results) {
            await WriteTextResultAsync(writer, result).ConfigureAwait(false);
        }
    }

    private static async Task WriteTextResultAsync(TextWriter writer, ProvenanceResultDto result) {
        await writer.WriteLineAsync(result.Status + " | " + result.Operation + " | " + result.Summary).ConfigureAwait(false);
        await writer.WriteLineAsync("Owner: " + result.OwnerPackage).ConfigureAwait(false);
        if (result.OutputPath is not null) await writer.WriteLineAsync("Output: " + result.OutputPath).ConfigureAwait(false);
        ProvenanceReportDto? report = result.Inspection ?? result.Assessment?.Structural ?? result.After ?? result.Before;
        if (report is not null) {
            await writer.WriteLineAsync("Format: " + report.Format + "; carriers: " + report.Evidence.Count).ConfigureAwait(false);
            foreach (ProvenanceEvidenceDto evidence in report.Evidence) {
                await writer.WriteLineAsync("  " + evidence.Carrier + " | " + evidence.Location + " | valid=" + evidence.IsStructurallyValid.ToString().ToLowerInvariant()).ConfigureAwait(false);
            }
        }
        if (result.Assessment is not null) {
            if (result.Assessment.Verification is not null) {
                ProvenanceVerificationDto verification = result.Assessment.Verification;
                await writer.WriteLineAsync("Verification: " + verification.ProviderName + " | status=" + verification.Status).ConfigureAwait(false);
                foreach (string finding in verification.Findings) {
                    await writer.WriteLineAsync("  Verification finding: " + finding).ConfigureAwait(false);
                }
            }
            IReadOnlyList<ProvenanceTextFindingDto> textFindings = result.Assessment.TextIntegrity ?? Array.Empty<ProvenanceTextFindingDto>();
            await writer.WriteLineAsync("Text integrity: " + textFindings.Count + " finding(s)").ConfigureAwait(false);
            foreach (ProvenanceTextFindingDto finding in textFindings) {
                await writer.WriteLineAsync(
                    "  " + finding.Risk + " | " + finding.Kind + " | " + finding.UnicodeNotation +
                    " | offset=" + finding.TextOffset + " | " + finding.Location).ConfigureAwait(false);
            }
            foreach (ProvenanceSignalDto signal in result.Assessment.ProviderSignals) {
                await writer.WriteLineAsync(
                    "Provider signal: " + signal.ProviderName + " | " + signal.SignalKind + " | status=" + signal.Status).ConfigureAwait(false);
                foreach (string finding in signal.Findings) {
                    await writer.WriteLineAsync("  Provider finding: " + finding).ConfigureAwait(false);
                }
            }
        }
    }

    private static ProvenanceCapabilityDto ToDto(OfficeProvenanceWorkflowCapability capability) => new(
        capability.Id,
        capability.Label,
        capability.Extensions,
        capability.OwnerPackage,
        capability.CanInspect,
        capability.CanAssess,
        capability.CanRemove,
        capability.Notes);

    private static ProvenanceResultDto ToDto(OfficeProvenanceWorkflowResult result) => new(
        "officeimo.provenance.result.v1",
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

internal sealed record ProvenanceCapabilitiesDto(string Schema, IReadOnlyList<ProvenanceCapabilityDto> Capabilities);
internal sealed record ProvenanceCapabilityDto(string Id, string Label, IReadOnlyList<string> Extensions, string OwnerPackage, bool CanInspect, bool CanAssess, bool CanRemove, string Notes);
internal sealed record ProvenanceBatchDto(string Schema, IReadOnlyList<ProvenanceResultDto> Results);
internal sealed record ProvenanceResultDto(
    string Schema,
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
internal sealed record ProvenanceReportDto(
    string Format,
    IReadOnlyList<ProvenanceEvidenceDto> Evidence,
    IReadOnlyList<string> Diagnostics,
    bool HasC2paManifest,
    bool HasExternalC2paManifest,
    bool HasGenerativeAiDeclaration);
internal sealed record ProvenanceEvidenceDto(string Carrier, string Location, bool IsStructurallyValid, long PayloadLength, string? Value, string DigitalSourceKind);
internal sealed record ProvenanceAssessmentDto(ProvenanceReportDto Structural, ProvenanceVerificationDto? Verification, IReadOnlyList<ProvenanceTextFindingDto>? TextIntegrity, IReadOnlyList<ProvenanceSignalDto> ProviderSignals);
internal sealed record ProvenanceVerificationDto(string ProviderName, string Status, IReadOnlyList<string> Findings, string? RawReport);
internal sealed record ProvenanceTextFindingDto(string Kind, string Risk, int TextOffset, int TextLength, string UnicodeNotation, string Location);
internal sealed record ProvenanceSignalDto(string ProviderName, string SignalKind, string Status, IReadOnlyList<string> Findings);
internal sealed record ProvenanceChangeDto(string Carrier, string Location, long RemovedBytes);
internal sealed record ProvenanceDiagnosticDto(string Code, string Message, string Severity, string? Stage, IReadOnlyDictionary<string, string> Details);

[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(ProvenanceCapabilitiesDto))]
[JsonSerializable(typeof(ProvenanceResultDto))]
[JsonSerializable(typeof(ProvenanceBatchDto))]
internal sealed partial class ProvenanceJsonContext : JsonSerializerContext;

using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Provenance;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Provenance;

internal static class ProvenanceOutput {
    internal static async Task WriteCapabilitiesAsync(TextWriter writer, ProvenanceOutputFormat format) {
        ProvenanceCapabilitiesDto dto = new(
            "officeimo.provenance.capabilities.v2",
            OfficeProvenanceWorkflowCatalog.All.Select(ToDto).ToArray());
        if (format == ProvenanceOutputFormat.Json) {
            await writer.WriteLineAsync(JsonSerializer.Serialize(dto, ProvenanceJsonContext.Default.ProvenanceCapabilitiesDto)).ConfigureAwait(false);
            return;
        }
        foreach (ProvenanceCapabilityDto capability in dto.Capabilities) {
            await writer.WriteLineAsync(
                capability.Id + " | " + capability.OwnerPackage + " | " +
                string.Join(',', capability.Extensions) + " | remove=" + capability.CanRemove.ToString().ToLowerInvariant() +
                " | browser=" + capability.BrowserAvailable.ToString().ToLowerInvariant()).ConfigureAwait(false);
        }
    }

    internal static async Task WriteResultAsync(
        TextWriter writer,
        OfficeProvenanceWorkflowResult result,
        ProvenanceOutputFormat format) {
        ProvenanceResultDto dto = OfficeProvenanceReportSerializer.Create(result);
        if (format == ProvenanceOutputFormat.Json) {
            await writer.WriteLineAsync(OfficeProvenanceReportSerializer.Serialize(result)).ConfigureAwait(false);
            return;
        }
        await WriteTextResultAsync(writer, dto).ConfigureAwait(false);
    }

    internal static async Task WriteBatchAsync(
        TextWriter writer,
        IReadOnlyList<OfficeProvenanceWorkflowResult> results,
        ProvenanceOutputFormat format) {
        if (format == ProvenanceOutputFormat.Json) {
            await writer.WriteLineAsync(OfficeProvenanceReportSerializer.SerializeBatch(results)).ConfigureAwait(false);
            return;
        }
        foreach (ProvenanceResultDto result in results.Select(OfficeProvenanceReportSerializer.Create)) {
            await WriteTextResultAsync(writer, result).ConfigureAwait(false);
        }
    }

    private static async Task WriteTextResultAsync(TextWriter writer, ProvenanceResultDto result) {
        await writer.WriteLineAsync(result.Status + " | " + result.Operation + " | " + result.Summary).ConfigureAwait(false);
        await writer.WriteLineAsync("Owner: " + result.OwnerPackage).ConfigureAwait(false);
        await writer.WriteLineAsync("Checks: structural=" + result.Checks.Structural + "; text integrity=" + result.Checks.TextIntegrity +
            "; verification=" + result.Checks.Verification + "; provider signals=" + result.Checks.ProviderSignals).ConfigureAwait(false);
        if (result.OutputPath is not null) await writer.WriteLineAsync("Output: " + result.OutputPath).ConfigureAwait(false);
        ProvenanceReportDto? report = result.Inspection ?? result.Assessment?.Structural ?? result.After ?? result.Before;
        if (report is not null) {
            await writer.WriteLineAsync("Format: " + report.Format + "; carriers: " + report.Evidence.Count).ConfigureAwait(false);
            foreach (ProvenanceEvidenceDto evidence in report.Evidence) {
                await writer.WriteLineAsync("  " + evidence.Carrier + " | " + evidence.Location + " | valid=" + evidence.IsStructurallyValid.ToString().ToLowerInvariant()).ConfigureAwait(false);
            }
        }
        if (result.Assessment is not null) {
            await writer.WriteLineAsync("Verification check: " + result.Assessment.VerificationStatus + "; provider checks: " + result.Assessment.ProviderSignalsStatus).ConfigureAwait(false);
            if (result.Assessment.Verification is not null) {
                ProvenanceVerificationDto verification = result.Assessment.Verification;
                await writer.WriteLineAsync("Verification: " + verification.ProviderName + " | status=" + verification.Status).ConfigureAwait(false);
                foreach (string finding in verification.Findings) {
                    await writer.WriteLineAsync("  Verification finding: " + finding).ConfigureAwait(false);
                }
            }
            IReadOnlyList<ProvenanceTextFindingDto> textFindings = result.Assessment.TextIntegrity ?? Array.Empty<ProvenanceTextFindingDto>();
            await writer.WriteLineAsync("Text integrity: " + result.Assessment.TextIntegrityStatus +
                (result.Assessment.TextIntegrity is null ? "" : " | " + textFindings.Count + " finding(s)")).ConfigureAwait(false);
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
        foreach (ProvenanceDiagnosticDto diagnostic in result.Diagnostics) {
            string stage = diagnostic.Stage is null ? string.Empty : " | stage=" + diagnostic.Stage;
            await writer.WriteLineAsync(
                "Diagnostic: " + diagnostic.Severity + " | " + diagnostic.Code + stage + " | " + diagnostic.Message).ConfigureAwait(false);
            foreach (KeyValuePair<string, string> detail in diagnostic.Details.OrderBy(
                         static pair => pair.Key,
                         StringComparer.Ordinal)) {
                await writer.WriteLineAsync("  " + detail.Key + ": " + detail.Value).ConfigureAwait(false);
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
        capability.MemoryOnlyExtensions,
        capability.BrowserAvailable,
        capability.BrowserLabel,
        capability.Formats.Select(static format => new ProvenanceCapabilityFormatDto(
            format.Extension,
            format.AssetFormats.Select(static assetFormat => assetFormat.ToString()).ToArray(),
            format.MemoryOnlyAvailable,
            format.BrowserAvailable)).ToArray(),
        capability.Notes);


}

internal sealed record ProvenanceCapabilitiesDto(string Schema, IReadOnlyList<ProvenanceCapabilityDto> Capabilities);
internal sealed record ProvenanceCapabilityDto(
    string Id,
    string Label,
    IReadOnlyList<string> Extensions,
    string OwnerPackage,
    bool CanInspect,
    bool CanAssess,
    bool CanRemove,
    IReadOnlyList<string> MemoryOnlyExtensions,
    bool BrowserAvailable,
    string? BrowserLabel,
    IReadOnlyList<ProvenanceCapabilityFormatDto> Formats,
    string Notes);
internal sealed record ProvenanceCapabilityFormatDto(
    string Extension,
    IReadOnlyList<string> AssetFormats,
    bool MemoryOnlyAvailable,
    bool BrowserAvailable);
[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(ProvenanceCapabilitiesDto))]
internal sealed partial class ProvenanceJsonContext : JsonSerializerContext;

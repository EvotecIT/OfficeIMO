using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows;

/// <summary>Canonical versioned hidden-data report transport, shared by the CLI, the MCP tools and the browser.</summary>
public static class OfficeHiddenDataReportSerializer {
    /// <summary>The schema identifier carried by every serialized report.</summary>
    public const string Schema = "officeimo.hidden-data.result.v1";

    /// <summary>Serializes a report as camel-case JSON with string enums.</summary>
    /// <param name="report">The inspection or removal report.</param>
    /// <param name="sourceName">The inspected file name or path.</param>
    /// <param name="outputName">The cleaned copy's name or path, when the report comes from removal.</param>
    public static string Serialize(OfficeHiddenDataReport report, string sourceName, string? outputName = null) =>
        JsonSerializer.Serialize(Create(report, sourceName, outputName), HiddenDataReportJsonContext.Default.HiddenDataResultDto);

    /// <summary>Creates the transport document for a report.</summary>
    public static HiddenDataResultDto Create(OfficeHiddenDataReport report, string sourceName, string? outputName = null) {
        ArgumentNullException.ThrowIfNull(report);
        return new HiddenDataResultDto(Schema, outputName == null ? "inspect" : "remove", report.Format, sourceName, outputName, report.Changed,
            report.Findings.Select(finding => new HiddenDataFindingDto(
                finding.Kind.ToString(), finding.Outcome.ToString(), finding.Scope.ToString(), finding.Reason == OfficeHiddenDataReason.None ? null : finding.Reason.ToString(),
                finding.Location, finding.Part, finding.Description, finding.Value, finding.Note, finding.Provenance?.Carrier.ToString(),
                finding.Provenance?.Carrier == OfficeProvenanceCarrierKind.IptcDigitalSourceType ? finding.Provenance.DigitalSourceKind.ToString() : null))
            .ToArray());
    }
}

#pragma warning disable CS1591
public sealed record HiddenDataResultDto(string Schema, string Operation, string Format, string Source, string? Output, bool Changed,
    IReadOnlyList<HiddenDataFindingDto> Findings);

public sealed record HiddenDataFindingDto(string Kind, string Outcome, string Scope, string? Reason, string Location, string? Part,
    string Description, string? Value, string? Note, string? ProvenanceCarrier, string? DigitalSourceKind);
#pragma warning restore CS1591

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull, GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(HiddenDataResultDto))]
internal sealed partial class HiddenDataReportJsonContext : JsonSerializerContext;

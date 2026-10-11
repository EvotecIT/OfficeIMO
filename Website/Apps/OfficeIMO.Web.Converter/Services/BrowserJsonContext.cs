using System.Text.Json.Serialization;

namespace OfficeIMO.Web.Converter.Services;

// Report documents written by the browser tools. Named records with source-generated metadata keep
// serialization intact when the engine is fully trimmed; property names match the published JSON.

internal sealed record ConversionManifestDocument(
    string SchemaVersion,
    string ConversionId,
    string FidelityStatus,
    ManifestSource Source,
    ManifestOutput Output,
    ManifestEngine Engine,
    ManifestFontPack FontPack,
    ManifestPolicy Policy,
    ManifestLimits Limits,
    ManifestPerformance Performance,
    ManifestWarning[] Warnings);

internal sealed record ManifestSource(string FileName, long ByteCount, string Sha256);

internal sealed record ManifestOutput(string FileName, long ByteCount, string Sha256, int? PageCount, bool Tagged);

internal sealed record ManifestEngine(string Converter, string? Assembly, string Version, string? SourceCommit, string OptionProfile, ManifestProfile Profile);

internal sealed record ManifestProfile(string Id, string Label, string Description);

internal sealed record ManifestFontPack(string Id, string Fingerprint, string DefaultFamily, IReadOnlyList<string> Coverage, ManifestSubstitution[] Substitutions);

internal sealed record ManifestSubstitution(string Source, string Target, string Impact);

internal sealed record ManifestPolicy(string Resources, string TaggedStructure, bool SystemFonts, bool ExternalResources);

internal sealed record ManifestLimits(string Profile, long PackageBytes, int PackagePartCount, long PartUncompressedBytes, long TotalUncompressedBytes, double CompressionRatio);

internal sealed record ManifestPerformance(
    long ConversionMilliseconds,
    long PeakRetainedPageContentBytes,
    long PeakRetainedObjectBytes,
    long PeakRetainedCompletedPayloadBytes,
    bool PageContentSpilled,
    bool ObjectBufferSpilled,
    bool FinalArtifactBuffered,
    bool IsForwardOnlyObjectSerialization,
    long LargestSerializedObjectBytes,
    bool IsForwardOnlyLayout);

internal sealed record ManifestWarning(
    string Converter,
    string Code,
    string Source,
    string Message,
    string Severity,
    string Construct,
    int? PageNumber,
    bool CanChangePagination,
    Dictionary<string, string> Details);

internal sealed record SupportSummaryDocument(
    string SchemaVersion,
    SupportPrivacy Privacy,
    SupportSource Source,
    SupportOutput Output,
    ManifestProfile? Profile,
    SupportEngine Engine,
    SupportPerformance Performance,
    SupportWarning[] Warnings);

internal sealed record SupportPrivacy(bool IncludesDocumentContent, string DefaultPolicy);

internal sealed record SupportSource(string Extension, long ByteCount, string Sha256);

internal sealed record SupportOutput(
    long ByteCount,
    string Sha256,
    [property: JsonPropertyName("PageCount")] int? PageCount,
    bool Tagged);

internal sealed record SupportEngine(string? Assembly, string Version, string FontPackId, string FontPackFingerprint);

internal sealed record SupportPerformance(
    [property: JsonPropertyName("ConversionMilliseconds")] long ConversionMilliseconds,
    [property: JsonPropertyName("PeakRetainedMemoryBytes")] long? PeakRetainedMemoryBytes);

internal sealed record SupportWarning(
    [property: JsonPropertyName("Code")] string Code,
    [property: JsonPropertyName("Source")] string Source,
    [property: JsonPropertyName("Message")] string Message,
    [property: JsonPropertyName("Severity")] string Severity,
    [property: JsonPropertyName("Construct")] string Construct,
    [property: JsonPropertyName("PageNumber")] int? PageNumber,
    [property: JsonPropertyName("CanChangePagination")] bool CanChangePagination);

internal sealed record PdfInspectionDocument(
    int SchemaVersion,
    string Tool,
    string Engine,
    bool BrowserLocal,
    PdfInspectionSource Source,
    string Summary,
    Dictionary<string, string> Details,
    PdfInspectionMessage[] Messages);

internal sealed record PdfInspectionSource(string FileName, long Bytes);

internal sealed record PdfInspectionMessage(string Title, string Message);

internal sealed record FontPackManifest(
    string Id,
    IReadOnlyList<string> Coverage,
    IReadOnlyList<FontPackFont> Fonts,
    IReadOnlyList<FontPackSubstitution> Substitutions);

internal sealed record FontPackFont(string Family, IReadOnlyList<FontPackFile>? Files);

internal sealed record FontPackFile(string Name, string Sha256);

internal sealed record FontPackSubstitution(string Source, string Target, string Impact);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, WriteIndented = true)]
[JsonSerializable(typeof(ConversionManifestDocument))]
[JsonSerializable(typeof(SupportSummaryDocument))]
[JsonSerializable(typeof(PdfInspectionDocument))]
internal sealed partial class BrowserReportJsonContext : JsonSerializerContext;

[JsonSourceGenerationOptions(PropertyNameCaseInsensitive = true)]
[JsonSerializable(typeof(FontPackManifest))]
internal sealed partial class BrowserFontPackJsonContext : JsonSerializerContext;

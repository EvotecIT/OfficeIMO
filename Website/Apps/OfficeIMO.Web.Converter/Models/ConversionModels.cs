namespace OfficeIMO.Web.Converter.Models;

public enum ConversionInputKind {
    File,
    Text
}

public sealed record ConversionRoute(
    string Id,
    string Source,
    string Target,
    string Title,
    string Description,
    ConversionInputKind InputKind,
    string Accept,
    string EnginePath,
    string AccentClass,
    string OutputModel,
    string SupportLevel,
    string SupportLabel,
    string SupportEvidence,
    string KnownLimitations);

public sealed record SelectedDocument(string Name, string Extension, string FormatLabel, long Size, byte[] Bytes);

public sealed record SampleDocument(string Label, string Path, string FileName, string Extension);

public sealed record ConversionDiagnostic(string Title, string Message, string ToneClass);

public enum BrowserPdfProfileKind {
    Faithful,
    Portable,
    Archival,
    Accessible,
    Diagnostic
}

public sealed record BrowserPdfProfile(
    BrowserPdfProfileKind Kind,
    string Id,
    string Label,
    string Description);

public sealed record BrowserPowerPointImportProfile(
    OfficeIMO.PowerPoint.Pdf.PdfPowerPointImportMode Mode,
    string Id,
    string Label,
    string Description,
    string FidelityStatus,
    string Projection);

public sealed record ConversionWarningView(
    string Code,
    string Source,
    string Message,
    string Severity,
    string Construct,
    int? PageNumber,
    bool CanChangePagination);

public sealed record ConversionResult(
    byte[] Bytes,
    string FileName,
    string ContentType,
    string? Text,
    string? HtmlPreview,
    IReadOnlyList<string> Warnings) {
    public string? FidelityStatus { get; init; }
    public string? ProvenanceSummary { get; init; }
    public BrowserConversionArtifact? CompanionReport { get; init; }
    public BrowserConversionArtifact? DebugOverlay { get; init; }
    public IReadOnlyList<ConversionWarningView> StructuredWarnings { get; init; } = [];
    public long? PeakRetainedMemoryBytes { get; init; }
    public int? PageCount { get; init; }
    public long ConversionMilliseconds { get; init; }
    public BrowserPdfProfile? Profile { get; init; }
    public SelectedDocument? SourceSnapshot { get; init; }
}

public sealed record WebMcpConversionResult(
    bool Success,
    string Route,
    string SourceFormat,
    string TargetFormat,
    string? OutputFileName,
    long OutputBytes,
    int WarningCount,
    long ElapsedMilliseconds,
    string Message);

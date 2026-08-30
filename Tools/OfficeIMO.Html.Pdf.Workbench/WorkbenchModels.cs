using System.Text.Json.Serialization;

namespace OfficeIMO.Html.Pdf.Workbench;

public enum HtmlPdfWorkbenchEngine {
    Managed,
    Chromium
}

public sealed class HtmlPdfWorkbenchSettings {
    public string PageSize { get; set; } = "A4";
    public bool Landscape { get; set; }
    public double MarginMillimeters { get; set; } = 12;
    public bool HonorCssPageSize { get; set; } = true;
    public bool PrintBackground { get; set; } = true;
    public bool TaggedPdf { get; set; } = true;
    public string Language { get; set; } = "en-US";
    public bool InteractiveForms { get; set; } = true;
    public bool Outline { get; set; } = true;
    public bool StrictFidelity { get; set; }

    public HtmlPdfWorkbenchSettings Clone() => (HtmlPdfWorkbenchSettings)MemberwiseClone();
}

public sealed record HtmlPdfWorkbenchRequest(
    string Html,
    string Css,
    HtmlPdfWorkbenchEngine Engine,
    HtmlPdfWorkbenchSettings Settings);

public sealed record HtmlPdfWorkbenchDiagnostic(
    string Severity,
    string Code,
    string Source,
    string Message);

public sealed record HtmlPdfWorkbenchEvidence(
    string Schema,
    DateTimeOffset CreatedUtc,
    string Engine,
    string RendererVersion,
    string InputSha256,
    string OutputSha256,
    long ElapsedMilliseconds,
    int PdfBytes,
    int PageCount,
    bool HasLoss,
    HtmlPdfWorkbenchSettings Options,
    IReadOnlyList<HtmlPdfWorkbenchDiagnostic> Diagnostics,
    BrowserCaptureEvidence? Browser);

public sealed record BrowserCaptureEvidence(
    string BrowserVersion,
    bool BrowserReused,
    bool Retried,
    int BlockedRequestCount,
    long QueueMilliseconds,
    long NavigationMilliseconds,
    long ReadinessMilliseconds,
    long PdfMilliseconds);

public sealed record HtmlPdfWorkbenchResult(
    byte[] PdfBytes,
    byte[] EvidenceBytes,
    HtmlPdfWorkbenchEvidence Evidence);

public sealed record HtmlPdfWorkbenchTemplate(string Id, string Name, string Description, string Html, string Css);

public sealed record WorkbenchArtifactLink(string Token, string PdfUrl, string EvidenceUrl);

public sealed record WorkbenchArtifact(byte[] PdfBytes, byte[] EvidenceBytes, DateTimeOffset CreatedUtc);

[JsonSourceGenerationOptions(WriteIndented = true, PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(HtmlPdfWorkbenchEvidence))]
internal partial class WorkbenchJsonContext : JsonSerializerContext;

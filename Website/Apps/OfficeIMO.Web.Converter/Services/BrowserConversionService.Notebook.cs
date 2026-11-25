using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Drawing;
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Fonts.Notebook;

namespace OfficeIMO.Web.Converter.Services;

internal sealed record BrowserNotebookSettings(bool HideCode, bool HideOutputs, string Title, string Author);

public sealed partial class BrowserConversionService {
    private static readonly Lazy<OfficeRenderingProfile> NotebookFonts = new(() =>
        new OfficeRenderingProfile("browser-notebook",
            new OfficeFontFaceCollection().Add("math", NotebookFontAssets.MathBytes)));
    private static readonly ManifestFontAsset[] NotebookFontManifest = [
        new("STIX Two Math", NotebookFontAssets.MathSha256),
        new("Liberation Mono", NotebookFontAssets.MonospaceSha256)
    ];

    private static PdfConversionPayload ConvertNotebookToPdfBytes(SelectedDocument file, BrowserPdfProfile profile,
        BrowserNotebookSettings? settings) {
        settings ??= new BrowserNotebookSettings(false, false, string.Empty, string.Empty);
        var load = NotebookLoadOptions.UntrustedDefaults;
        load.MaxInputBytes = MaxPackageBytes;
        using NotebookDocument notebook = NotebookDocument.Load(file.Bytes, load);
        PdfCreateOptions pdf = BrowserPortablePdfProfile.CreateOptions(profile);
        pdf.UseRenderingProfile(NotebookFonts.Value, OfficeRenderingProfileApplyMode.Overlay);
        pdf.EmbedStandardFontSnapshot(PdfStandardFont.Courier, NotebookFontAssets.MonospaceBytes, "LiberationMono");
        pdf.MaxGeneratedOutputBytes = BrowserPdfPolicy.MaxOutputBytes;
        var options = new NotebookPdfExportOptions {
            IncludeCode = !settings.HideCode,
            IncludeOutputs = !settings.HideOutputs,
            Title = settings.Title,
            Author = settings.Author,
            MarkdownOptions = new MarkdownPdfExportOptions {
                ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
                PdfOptions = pdf
            }
        };
        PdfDocumentConversionResult conversion = notebook.ConvertToPdfWithEvidence(options, default);
        using PdfDocument document = conversion.Value;
        (byte[] bytes, PdfSerializationReport serialization) = SaveWithEvidence(conversion);
        PdfConversionReport findings = conversion.WithAdditionalWarnings(conversion.SourceConversionReports
            .SelectMany(static stage => stage.Diagnostics)
            .Select(static finding => new PdfConversionWarning(
                finding.Format ?? "OfficeIMO.Notebook", finding.Code, finding.Location ?? string.Empty, finding.Message,
                finding.Severity switch {
                    OfficeDiagnosticSeverity.Error => PdfConversionWarningSeverity.Error,
                    OfficeDiagnosticSeverity.Warning => PdfConversionWarningSeverity.Warning,
                    _ => PdfConversionWarningSeverity.Information
                }, finding.LossKind))).Report;
        var limits = new ManifestLimits("browser", MaxPackageBytes, OutputBytes: BrowserPdfPolicy.MaxOutputBytes);
        byte[] configuration = JsonSerializer.SerializeToUtf8Bytes(settings, BrowserReportJsonContext.Default.BrowserNotebookSettings);
        string optionsFingerprint = Convert.ToHexStringLower(SHA256.HashData(configuration));
        return new PdfConversionPayload(bytes, findings, serialization,
            "OfficeIMO.Notebook.Pdf", "portable-deterministic", limits, optionsFingerprint, NotebookFontManifest);
    }
}

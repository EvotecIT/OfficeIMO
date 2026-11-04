using OfficeIMO.Pdf;
using OfficeIMO.Publisher;
using OfficeIMO.Publisher.Pdf;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

public sealed partial class BrowserConversionService {
    private static PdfConversionPayload ConvertPublisherToPdfBytes(SelectedDocument file, BrowserPdfProfile profile) {
        var loadOptions = new PublisherLoadOptions {
            Limits = { MaxInputBytes = (int)MaxPackageBytes },
            MaxPages = 128,
            MaxTotalImageBytes = 32 * 1024 * 1024,
            Fonts = BrowserPortablePdfProfile.CreateDrawingFonts()
        };
        PublisherDocument publication = PublisherDocument.Load(file.Bytes, loadOptions);
        PdfCreateOptions pdfOptions = BrowserPortablePdfProfile.CreateOptions(profile);
        pdfOptions.MaxGeneratedOutputBytes = BrowserPdfPolicy.MaxOutputBytes;
        PdfDocumentConversionResult conversion = publication.ConvertToPdfWithEvidence(
            new PublisherPdfExportOptions { PdfOptions = pdfOptions }, default);
        using PdfDocument document = conversion.Value;

        // The browser report has one warning list. Include the source stage so a recovered page
        // with approximated or unassessed Publisher content is not reported as faithful.
        conversion = conversion.WithAdditionalWarnings(publication.ReadReport.Diagnostics.Select(static diagnostic =>
            new PdfConversionWarning(
                diagnostic.Format ?? "OfficeIMO.Publisher",
                diagnostic.Code,
                diagnostic.Location ?? string.Empty,
                diagnostic.Message,
                diagnostic.Severity switch {
                    OfficeDiagnosticSeverity.Error => PdfConversionWarningSeverity.Error,
                    OfficeDiagnosticSeverity.Warning => PdfConversionWarningSeverity.Warning,
                    _ => PdfConversionWarningSeverity.Information
                },
                diagnostic.LossKind,
                details: new Dictionary<string, string> { ["lossKind"] = diagnostic.LossKind.ToString() })));
        (byte[] bytes, PdfSerializationReport serialization) = SaveWithEvidence(conversion);
        var limits = new ManifestLimits("browser", loadOptions.Limits.MaxInputBytes,
            PublicationPages: loadOptions.MaxPages, PublicationImageBytes: loadOptions.MaxTotalImageBytes,
            OutputBytes: BrowserPdfPolicy.MaxOutputBytes);
        return new PdfConversionPayload(bytes, conversion.Report, serialization, "OfficeIMO.Publisher.Pdf", "source-page-size", limits);
    }
}

using OfficeIMO.Drawing;
using OfficeIMO.Email;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;

namespace OfficeIMO.ConversionConsistency;

internal static partial class NativeAdapters {
    private static NativeExport ExportOther(ConsistencyCase contract, string source, ConsistencySuite suite,
        OfficeRenderingProfile profile, CancellationToken token) {
        switch (contract.Format) {
            case "one": {
                var document = OneNoteSectionReader.Read(source);
                var imageOptions = Images<OneNotePageBatchRenderingOptions>(suite, profile);
                imageOptions.DefaultFont = new OfficeFontInfo(suite.FontFamily, 11D);
                var pdfOptions = new OneNoteVisualPdfOptions {
                    PdfOptions = new PdfOptions().UseRenderingProfile(profile),
                    PageRendering = imageOptions, RasterScale = suite.Dpi / 72D
                };
                return Capture(document.ToVisualPdfDocumentResult(pdfOptions, token),
                    format => document.ExportImages(format, imageOptions), token) with { PdfRoute = "native-raster" };
            }
            case "vsdx": {
                var document = VisioDocument.Load(source);
                var imageOptions = Images<VisioImageExportOptions>(suite, profile);
                var pdfOptions = new VisioToPdfOptions {
                    ProjectionOptions = new PdfProjectionOptions { PdfOptions = new PdfOptions().UseRenderingProfile(profile) }
                };
                return Capture(document.ToPdfDocumentResult(pdfOptions, token),
                    format => document.ExportImages(format, imageOptions), token) with { PdfRoute = "native-semantic-projection" };
            }
            case "eml": {
                byte[] reference = ReferencePdf(contract, source);
                var document = EmailDocument.Load(source);
                var options = Images<EmailImageExportOptions>(suite, profile);
                ConfigureHtml(options, contract, suite);
                options.IncludeMessageHeaders = false;
                return Capture(reference, format => document.ExportImages(format, options), new(), token)
                    with { PdfRoute = "external-reference" };
            }
            case "epub": {
                byte[] reference = ReferencePdf(contract, source);
                var document = EpubDocument.Load(source, new EpubReadOptions { IncludeRawHtml = true });
                var options = Images<EpubImageExportOptions>(suite, profile);
                ConfigureHtml(options, contract, suite);
                options.IncludeChapterTitle = false;
                return Capture(reference, format => document.ExportImages(format, options), new(), token)
                    with { PdfRoute = "external-reference" };
            }
            default: throw new NotSupportedException("No native adapter registered for " + contract.Format);
        }
    }

    private static byte[] ReferencePdf(ConsistencyCase contract, string source) {
        if (string.IsNullOrWhiteSpace(contract.ReferencePdf))
            throw new InvalidDataException(contract.Format + " has no matching native PDF export; provide referencePdf relative to the source directory and describe its provenance in evidence.");
        return File.ReadAllBytes(ArtifactPaths.Resolve(Path.GetDirectoryName(source)!, contract.ReferencePdf!));
    }

    private static void ConfigureHtml(HtmlRenderOptions options, ConsistencyCase contract, ConsistencySuite suite) {
        options.Mode = HtmlRenderMode.Paged;
        options.DefaultFontFamily = suite.FontFamily;
        options.Margins = HtmlRenderMargins.All(0);
        options.PageSize = new OfficePageSize(contract.Pages[0].Width / (double)suite.Dpi, contract.Pages[0].Height / (double)suite.Dpi);
    }
}

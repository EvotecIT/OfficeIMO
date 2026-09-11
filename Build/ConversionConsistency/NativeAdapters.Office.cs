using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.OpenDocument;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.ConversionConsistency;

internal static partial class NativeAdapters {
    private static NativeExport ExportOffice(ConsistencyCase contract, string source, ConsistencySuite suite,
        OfficeRenderingProfile profile, CancellationToken token) {
        switch (contract.Format) {
            case "docx": {
                using var document = WordDocument.Load(source);
                var imageOptions = Images<WordImageExportOptions>(suite, profile);
                var pdfOptions = new WordToPdfOptions().UseRenderingProfile(profile);
                return Capture(document.ToPdfDocumentResult(pdfOptions, token),
                    format => Collect(consumer => document.ExportImages(format, consumer, imageOptions, token)), token);
            }
            case "xlsx": {
                using var document = ExcelDocument.Load(source);
                var imageOptions = Images<ExcelWorkbookImageExportOptions>(suite, profile);
                var pdfOptions = new ExcelToPdfOptions().UseRenderingProfile(profile);
                return Capture(document.ToPdfDocumentResult(pdfOptions, token),
                    format => Collect(consumer => document.ExportImages(format, consumer, imageOptions, token)), token);
            }
            case "pptx": {
                using var document = PowerPointPresentation.Load(source);
                var imageOptions = Images<PowerPointPresentationImageExportOptions>(suite, profile);
                var pdfOptions = new PowerPointToPdfOptions().UseRenderingProfile(profile);
                return Capture(document.ToPdfDocumentResult(pdfOptions, token),
                    format => Collect(consumer => document.ExportImages(format, consumer, imageOptions, token)), token);
            }
            case "odt": {
                var document = OdtDocument.Load(source);
                var imageOptions = Images<WordImageExportOptions>(suite, profile);
                var pdfOptions = new WordToPdfOptions().UseRenderingProfile(profile);
                return Capture(document.ToPdfDocumentResult(pdfOptions: pdfOptions, cancellationToken: token),
                    format => document.ExportImages(format, imageOptions, cancellationToken: token), token);
            }
            case "ods": {
                var document = OdsDocument.Load(source);
                var imageOptions = Images<ExcelWorkbookImageExportOptions>(suite, profile);
                var pdfOptions = new ExcelToPdfOptions().UseRenderingProfile(profile);
                return Capture(document.ToPdfDocumentResult(pdfOptions: pdfOptions, cancellationToken: token),
                    format => document.ExportImages(format, imageOptions, cancellationToken: token), token);
            }
            case "odp": {
                var document = OdpPresentation.Load(source);
                var imageOptions = Images<PowerPointPresentationImageExportOptions>(suite, profile);
                var pdfOptions = new PowerPointToPdfOptions().UseRenderingProfile(profile);
                return Capture(document.ToPdfDocumentResult(pdfOptions: pdfOptions, cancellationToken: token),
                    format => document.ExportImages(format, imageOptions, cancellationToken: token), token);
            }
            case "pdf": {
                byte[] bytes = File.ReadAllBytes(source);
                var document = PdfDocument.Load(bytes);
                var imageOptions = Images<PdfImageExportOptions>(suite, profile);
                return Capture(bytes, format => document.Render.ExportImages(format, imageOptions), new(), token);
            }
            default: return ExportOther(contract, source, suite, profile, token);
        }
    }

    private static IReadOnlyList<OfficeImageExportResult> Collect(Action<OfficeImageExportConsumer> export) {
        var images = new List<OfficeImageExportResult>();
        export(images.Add);
        return images;
    }

    private static T Images<T>(ConsistencySuite suite, OfficeRenderingProfile profile) where T : OfficeImageExportOptions, new() {
        var options = new T { TargetDpi = suite.Dpi, BackgroundColor = OfficeColor.White };
        options.UseRenderingProfile(profile);
        return options;
    }

    // Materialize while the native document is open; do not retain a delegate over a disposed package.
    private static NativeExport Capture(PdfDocumentConversionResult pdf,
        Func<OfficeImageExportFormat, IReadOnlyList<OfficeImageExportResult>> images, CancellationToken token) =>
        Capture(pdf.ToBytes(token), images, pdf.Report.Warnings.Select(item => item.Code).ToList(), token);

    private static NativeExport Capture(byte[] pdf, Func<OfficeImageExportFormat, IReadOnlyList<OfficeImageExportResult>> images,
        List<string> diagnostics, CancellationToken token) {
        var results = new Dictionary<OfficeImageExportFormat, IReadOnlyList<OfficeImageExportResult>>();
        foreach (OfficeImageExportFormat format in Enum.GetValues<OfficeImageExportFormat>()) {
            token.ThrowIfCancellationRequested();
            results.Add(format, images(format));
        }
        return new NativeExport(pdf, format => results[format], diagnostics);
    }
}

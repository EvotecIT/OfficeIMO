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
                var pdfOptions = new WordPdfExportOptions().UseRenderingProfile(profile);
                using PdfDocument pdf = document.ToPdf(pdfOptions, out OfficeConversionReport report, token);
                return Capture(pdf, report,
                    format => Collect(consumer => document.ExportImages(format, consumer, imageOptions, token)), token);
            }
            case "xlsx": {
                using var document = ExcelDocument.Load(source);
                var imageOptions = Images<ExcelImageExportOptions>(suite, profile);
                var pdfOptions = new ExcelPdfExportOptions().UseRenderingProfile(profile);
                using PdfDocument pdf = document.ToPdf(pdfOptions, out OfficeConversionReport report, token);
                return Capture(pdf, report,
                    format => Collect(consumer => document.ExportImages(format, consumer, imageOptions, token)), token);
            }
            case "pptx": {
                using var document = PowerPointDocument.Load(source);
                var imageOptions = Images<PowerPointImageExportOptions>(suite, profile);
                var pdfOptions = new PowerPointPdfExportOptions().UseRenderingProfile(profile);
                using PdfDocument pdf = document.ToPdf(pdfOptions, out OfficeConversionReport report, token);
                return Capture(pdf, report,
                    format => Collect(consumer => document.ExportImages(format, consumer, imageOptions, token)), token);
            }
            case "odt": {
                var document = OdtDocument.Load(source);
                var imageOptions = Images<WordImageExportOptions>(suite, profile);
                var pdfOptions = new WordPdfExportOptions().UseRenderingProfile(profile);
                return Capture(document.ToPdfDocumentResult(pdfOptions: pdfOptions, cancellationToken: token),
                    format => document.ExportImages(format, imageOptions, cancellationToken: token), token);
            }
            case "ods": {
                var document = OdsDocument.Load(source);
                var imageOptions = Images<ExcelImageExportOptions>(suite, profile);
                var pdfOptions = new ExcelPdfExportOptions().UseRenderingProfile(profile);
                return Capture(document.ToPdfDocumentResult(pdfOptions: pdfOptions, cancellationToken: token),
                    format => document.ExportImages(format, imageOptions, cancellationToken: token), token);
            }
            case "odp": {
                var document = OdpPresentation.Load(source);
                var imageOptions = Images<PowerPointImageExportOptions>(suite, profile);
                var pdfOptions = new PowerPointPdfExportOptions().UseRenderingProfile(profile);
                return Capture(document.ToPdfDocumentResult(pdfOptions: pdfOptions, cancellationToken: token),
                    format => document.ExportImages(format, imageOptions, cancellationToken: token), token);
            }
            case "pdf": {
                byte[] bytes = File.ReadAllBytes(source);
                using var document = PdfDocument.Load(source);
                var imageOptions = Images<PdfImageExportOptions>(suite, profile);
                var diagnostics = new List<string>();
                var details = new List<string>();
                NativeExport exported = Capture(bytes, format => {
                    imageOptions.Format = format;
                    IReadOnlyList<byte[]> pages = document.ExportImages(imageOptions, out var report, token);
                    diagnostics.AddRange(report.Diagnostics.Select(item => item.Code));
                    details.AddRange(report.Diagnostics.Select(item => item.Code + ": " + item.Message + " [" + item.Location + "]"));
                    return pages.Select((page, index) => {
                        if (!OfficeImageReader.TryValidateContent(page, format.GetFileExtension(), out OfficeImageInfo info))
                            throw new InvalidDataException("The PDF image export returned an invalid " + format + " payload.");
                        return new OfficeImageExportResult(format, info.Width, info.Height, page, name: "Page " + (index + 1));
                    }).ToArray();
                }, diagnostics, token);
                return exported with { DiagnosticDetails = details };
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
    private static NativeExport Capture(PdfDocument pdf, OfficeConversionReport report,
        Func<OfficeImageExportFormat, IReadOnlyList<OfficeImageExportResult>> images, CancellationToken token) {
        byte[] bytes = pdf.ToBytes(new PdfSaveOptions { AllowLoss = true }, token);
        return Capture(bytes, images, report.Diagnostics.Select(item => item.Code).ToList(), token) with {
            DiagnosticDetails = report.Diagnostics.Select(item => item.Code + ": " + item.Message + " [" + item.Location + "]").ToList()
        };
    }

    private static NativeExport Capture(PdfDocumentConversionResult pdf,
        Func<OfficeImageExportFormat, IReadOnlyList<OfficeImageExportResult>> images, CancellationToken token) {
        byte[] bytes = pdf.ToBytes(token);
        return Capture(bytes, images, pdf.Report.Warnings.Select(item => item.Code).ToList(), token) with {
            DiagnosticDetails = pdf.Report.Warnings.Select(item => item.Code + ": " + item.Message + " [" + item.Source + "]").ToList()
        };
    }

    private static NativeExport Capture(byte[] pdf, Func<OfficeImageExportFormat, IReadOnlyList<OfficeImageExportResult>> images,
        List<string> diagnostics, CancellationToken token) {
        var results = new Dictionary<OfficeImageExportFormat, IReadOnlyList<OfficeImageExportResult>>();
        foreach (OfficeImageExportFormat format in ComparedImageFormats.All) {
            token.ThrowIfCancellationRequested();
            results.Add(format, images(format));
        }
        return new NativeExport(pdf, format => results[format], diagnostics);
    }
}

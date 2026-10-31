using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Word;
using OfficeIMO.Visio;
using OfficeIMO.Word;

namespace OfficeIMO;

public static partial class OfficeDocument {
    /// <summary>Inspects the current in-memory native model using default inspection options.</summary>
    public static OfficeDocumentInfo Inspect(object document) => Inspect(document, null, default);

    /// <summary>Inspects the current in-memory native model using the supplied inspection options.</summary>
    public static OfficeDocumentInfo Inspect(object document, OfficeInspectOptions? options) => Inspect(document, options, default);

    /// <summary>Inspects a live native model without opening its associated path, saving it or disposing caller-owned state.</summary>
    /// <remarks>Counts describe the current in-memory model, including unsaved edits. Optional native details are detached snapshots.</remarks>
    public static OfficeDocumentInfo Inspect(object document, OfficeInspectOptions? options,
        CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        OfficeInspectOptions settings = options ?? new();
        if (settings.Reader == null || settings.Formats == null) throw new ArgumentException("Reader and format settings are required.", nameof(options));
        if (settings.ContentSafety != null || settings.Provenance != null || settings.PackageSignatures != null ||
            settings.VbaSignatures != null || settings.WordMacroSignatures != null || settings.PdfPreflight != null)
            throw new NotSupportedException("Selected source reports require the path or stream inspection overload.");
        var result = new OfficeDocumentInfo(document switch {
            WordDocument word => word.FilePath ?? string.Empty,
            ExcelDocument excel => excel.FilePath ?? string.Empty,
            PowerPointDocument presentation => presentation.FilePath ?? string.Empty,
            PdfDocument pdf => pdf.FilePath ?? string.Empty,
            VisioDocument visio => visio.FilePath ?? string.Empty,
            OfficeDocumentReadResult read => read.Source?.Path ?? string.Empty,
            _ => string.Empty
        });
        PopulateSummary(result, document, settings, cancellationToken);
        return result;
    }

    private static void PopulateSummary(OfficeDocumentInfo result, object document, OfficeInspectOptions settings,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (document) {
            case WordDocument word:
                result.Format = ReaderInputKind.Word;
                result.PageCount = word.Statistics?.Pages;
                if (settings.Detailed) result.Details = word.Inspect(includeImageBytes: false, cancellationToken);
                if (settings.IncludeContent) result.Content = word.ReadDocument(result.Source, settings.Reader, settings.Formats.Word, cancellationToken);
                break;
            case ExcelDocument excel:
                result.Format = ReaderInputKind.Excel;
                result.SheetCount = excel.Sheets.Count;
                if (settings.Detailed) result.Details = excel.Inspect(settings.ExcelInspection ?? settings.Formats.Excel?.ReadOptions, cancellationToken);
                if (settings.IncludeContent) result.Content = excel.ReadDocument(result.Source, settings.Reader, settings.Formats.Excel, cancellationToken);
                break;
            case PowerPointDocument presentation:
                result.Format = ReaderInputKind.PowerPoint;
                result.SlideCount = presentation.Slides.Count;
                if (settings.Detailed) result.Details = presentation.Inspect(settings.PowerPointInspection);
                if (settings.IncludeContent) result.Content = presentation.ReadDocument(result.Source, settings.Reader, settings.Formats.PowerPoint, cancellationToken);
                break;
            case PdfDocument pdf:
                ApplyPdfSummary(result, pdf.Preflight(settings.Formats.Pdf?.ReadOptions ?? settings.Formats.Pdf?.LoadOptions, cancellationToken), settings);
                if (settings.IncludeContent && result.PdfPreflight!.CanReadLogicalObjects)
                    result.Content = OfficeIMO.Reader.Pdf.PdfDocumentReaderExtensions.ReadDocument(pdf.Read(settings.Formats.Pdf?.ReadOptions, cancellationToken),
                        result.Source, settings.Reader, settings.Formats.Pdf, cancellationToken);
                break;
            case VisioDocument visio:
                result.Format = ReaderInputKind.Visio;
                result.PageCount = visio.Pages.Count;
                if (settings.Detailed) result.Details = visio.Inspect();
                if (settings.IncludeContent) result.Content = OfficeIMO.Reader.Visio.VisioReaderAdapter.ReadDocument(visio,
                    result.Source, settings.Reader, settings.Formats.Visio, cancellationToken);
                break;
            case MarkdownDocument markdown:
                result.Format = ReaderInputKind.Markdown;
                if (settings.Detailed) result.Details = markdown.ToMarkdown();
                if (settings.IncludeContent) {
                    using var source = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(markdown.ToMarkdown()));
                    var reader = settings.DocumentReader ?? new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers(settings.Formats).Build();
                    ApplyReadSummary(result, reader.ReadDocument(source, "document.md", settings.Reader, cancellationToken), settings);
                }
                break;
            case OfficeDocumentReadResult read:
                ApplyReadSummary(result, read, settings);
                break;
            default:
                throw new NotSupportedException("Live inspection is not available for " + document.GetType().FullName + ".");
        }
        cancellationToken.ThrowIfCancellationRequested();
    }
}

using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Canonical PDF page image-export entry points.</summary>
public static partial class PdfImageExportExtensions {
    /// <summary>Exports pages from an authored or opened PDF document using the shared five-format result contract.</summary>
    internal static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this PdfDocument document,
        OfficeImageExportFormat format,
        PdfImageExportOptions? options = null,
        PdfPageSelection? selection = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        return PdfImageExportEngine.Export(
            token => {
                token.ThrowIfCancellationRequested();
                PdfReadDocument readDocument = document.GetReadSnapshot(cancellationToken: token).Document;
                token.ThrowIfCancellationRequested();
                return readDocument;
            },
            format,
            options?.Clone() ?? new PdfImageExportOptions(),
            _ => selection,
            initialDiagnostics: null,
            cancellationToken: cancellationToken);
    }

    /// <summary>Exports one loaded PDF page using the shared five-format result contract.</summary>
    internal static OfficeImageExportResult ExportImage(
        this PdfReadPage page,
        OfficeImageExportFormat format,
        PdfImageExportOptions? options = null) =>
        page.ExportImage(format, options, CancellationToken.None);

    /// <summary>Exports one loaded PDF page and observes cancellation throughout supported render stages.</summary>
    internal static OfficeImageExportResult ExportImage(
        this PdfReadPage page,
        OfficeImageExportFormat format,
        PdfImageExportOptions? options,
        CancellationToken cancellationToken) =>
        PdfImageExportEngine.Export(
            page,
            format,
            options?.Clone() ?? new PdfImageExportOptions(),
            cancellationToken: cancellationToken);

    /// <summary>Exports selected loaded PDF pages using the shared five-format result contract.</summary>
    internal static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this PdfReadDocument document,
        OfficeImageExportFormat format,
        PdfImageExportOptions? options = null,
        PdfPageSelection? selection = null,
        CancellationToken cancellationToken = default) =>
        PdfImageExportEngine.Export(
            document,
            format,
            options?.Clone() ?? new PdfImageExportOptions(),
            selection,
            initialDiagnostics: null,
            cancellationToken: cancellationToken);

    /// <summary>
    /// Exports pages from a source-to-PDF conversion while preserving its conversion diagnostics.
    /// This is the shared paged-image bridge for Markdown, AsciiDoc, LaTeX, RTF, OneNote, and other PDF adapters.
    /// </summary>
    internal static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this PdfDocumentConversionResult conversion,
        OfficeImageExportFormat format,
        PdfImageExportOptions? options = null,
        PdfPageSelection? selection = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(conversion, nameof(conversion));
        return PdfImageExportEngine.Export(
            token => {
                token.ThrowIfCancellationRequested();
                PdfReadDocument document = PdfReadDocument.Open(
                    conversion.ToBytes(token),
                    options: null,
                    cancellationToken: token);
                token.ThrowIfCancellationRequested();
                return document;
            },
            format,
            options?.Clone() ?? new PdfImageExportOptions(),
            _ => selection,
            initialDiagnostics: null,
            () => PdfImageExportEngine.MapConversionDiagnostics(conversion),
            cancellationToken);
    }

}

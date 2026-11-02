using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word;

/// <summary>
/// Thin image-export adapter over the canonical Word page layout and PDF page scenes.
/// It owns selection and option mapping, never document measurement or pagination.
/// </summary>
internal static class WordPageImageExport {
    internal static int EstimatePageCount(WordDocument document) =>
        OpenPages(document, new WordImageExportOptions(), CancellationToken.None).Document.Pages.Count;

    internal static OfficeImageExportResult Render(WordDocument document,
        OfficeImageExportFormat format, WordImageExportOptions options, CancellationToken cancellationToken) {
        WordImageExportOptions single = options.Clone();
        single.PageCount = 1;
        single.PageSelection = null;
        return RenderPages(document, format, single, cancellationToken)[0];
    }

    internal static IReadOnlyList<OfficeImageExportResult> RenderPages(WordDocument document,
        OfficeImageExportFormat format, WordImageExportOptions options, CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        RenderPages(document, format, options, results.Add, cancellationToken);
        return results.AsReadOnly();
    }

    internal static void RenderPages(WordDocument document, OfficeImageExportFormat format,
        WordImageExportOptions options, OfficeImageExportConsumer consumer, CancellationToken cancellationToken = default) {
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        PdfCore.PdfDocumentConversionResult? conversion = null;
        PdfCore.PdfImageExportEngine.ExportEach(
            token => {
                var pages = OpenPages(document, options, token);
                conversion = pages.Conversion;
                return pages.Document;
            },
            format,
            options.ToPdfImageOptions(),
            pages => ResolveSelection(options, pages.Pages.Count),
            result => consumer(result.WithIdentity(result.Name, "Word document")),
            initialDiagnostics: null,
            () => PdfCore.PdfImageExportEngine.MapConversionDiagnostics(conversion!),
            cancellationToken);
    }

    internal static WordDocumentVisualInfo CreateSnapshot(WordDocument document, WordImageExportOptions options) {
        WordImageExportOptions single = options.Clone();
        single.PageCount = 1;
        single.PageSelection = null;
        return CreateSnapshots(document, single)[0];
    }

    internal static IReadOnlyList<WordDocumentVisualInfo> CreateSnapshots(WordDocument document,
        WordImageExportOptions options, CancellationToken cancellationToken = default) {
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(options.RenderTimeout, cancellationToken);
        try {
            var pages = OpenPages(document, options, execution.Token);
            IReadOnlyList<int> selection = ResolveSelection(options, pages.Document.Pages.Count).Resolve(pages.Document.Pages.Count);
            if (selection.Count > options.MaximumOutputCount)
                throw new InvalidOperationException("Word page snapshot count exceeds " + nameof(options.MaximumOutputCount) + ".");
            var snapshots = new List<WordDocumentVisualInfo>(selection.Count);
            foreach (int pageNumber in selection) {
                execution.ThrowIfCancellationRequested();
                PdfCore.PdfReadPage page = pages.Document.Pages[pageNumber - 1];
                OfficeDrawing drawing = page.ToDrawing(execution.Token, scene => scene.ApplyImageExportOptions(options),
                    options.GetEffectiveScale(page.GetPageSize().Width, page.GetPageSize().Height), options.MaximumRasterPixels);
                var diagnostics = new List<OfficeImageExportDiagnostic>(PdfCore.PdfImageExportEngine.MapConversionDiagnostics(pages.Conversion));
                foreach (PdfCore.PdfRenderCapabilityDiagnostic diagnostic in page.GetRenderCapabilityDiagnostics(cancellationToken: execution.Token))
                    diagnostics.Add(new OfficeImageExportDiagnostic(
                        diagnostic.SupportLevel == PdfCore.PdfRenderSupportLevel.Supported ? OfficeImageExportDiagnosticSeverity.Info : OfficeImageExportDiagnosticSeverity.Warning,
                        diagnostic.Code, diagnostic.Message, "Word page " + pageNumber, diagnostic.LossKind));
                drawing.AppendFontDiagnostics(diagnostics, "Word page " + pageNumber);
                snapshots.Add(new WordDocumentVisualInfo(drawing, pageNumber - 1, diagnostics.AsReadOnly(), page.GetTextSpans(includeArtifactText: true, execution.Token)));
            }
            execution.ThrowIfCancellationRequested();
            return snapshots.AsReadOnly();
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    private static (PdfCore.PdfReadDocument Document, PdfCore.PdfDocumentConversionResult Conversion) OpenPages(
        WordDocument document, WordImageExportOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        WordPdfExportOptions layout = options.LayoutOptions.Clone();
        if (options.Fonts.Faces.Count > 0 || options.TextShapingProvider != null || options.TextShapingLanguage != null || options.ImageCodec != null)
            layout.UseRenderingProfile(new OfficeRenderingProfile("Word image export", options.Fonts,
                options.TextShapingProvider, options.TextShapingLanguage, options.ImageCodec), OfficeRenderingProfileApplyMode.Overlay);
        PdfCore.PdfDocumentConversionResult conversion = document.ConvertToPdfWithEvidence(layout, cancellationToken);
        PdfCore.PdfReadDocument pages = PdfCore.PdfReadDocument.Open(conversion.ToBytes(cancellationToken),
            options: null, cancellationToken: cancellationToken);
        return (pages, conversion);
    }

    private static PdfCore.PdfPageSelection ResolveSelection(WordImageExportOptions options, int totalPages) {
        if (options.PageSelection != null) return options.PageSelection;
        if (options.PageIndex >= totalPages)
            throw new ArgumentOutOfRangeException(nameof(options.PageIndex), "Page index is outside the document.");
        int count = Math.Min(options.PageCount ?? totalPages, totalPages - options.PageIndex);
        return PdfCore.PdfPageSelection.FromRanges(new PdfCore.PdfPageRange(options.PageIndex + 1, options.PageIndex + count));
    }
}

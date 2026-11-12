using System.Diagnostics;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageImageRenderer {
    /// <summary>Renders all pages or a caller-ordered page selection with bounded per-page reports.</summary>
    public static IReadOnlyList<PdfPageRenderResult> RenderPages(
        byte[] pdf,
        PdfPageSelection? selection = null,
        PdfPageRenderOptions? options = null,
        PdfLoadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        return RenderPages(token => PdfReadDocument.Open(pdf, readOptions, token), selection, options, cancellationToken);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfPageRenderResult> RenderPagesCore(
        Func<CancellationToken, PdfReadDocument> getDocument,
        Func<int, int[]> resolvePages,
        PdfPageRenderOptions? options,
        CancellationToken cancellationToken) {
        // One detached snapshot decides extraction permission and produces the output, so changes the caller makes to
        // its options during the render cannot select SVG, a codec, a shaping provider, or another font afterwards.
        PdfPageRenderOptions effectiveOptions = (options ?? new PdfPageRenderOptions()).CreateEffectiveImageExportOptions<PdfPageRenderOptions>();
        effectiveOptions.Validate();
        bool pixelsOnly = ProducesPixelsOnly(effectiveOptions);
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            effectiveOptions.RenderTimeout,
            cancellationToken);
        try {
            execution.Token.ThrowIfCancellationRequested();
            PdfReadDocument document = getDocument(execution.Token);
            Guard.NotNull(document, nameof(document));
            execution.Token.ThrowIfCancellationRequested();
            int[] pages = resolvePages(document.Pages.Count);
            execution.Token.ThrowIfCancellationRequested();
            if (pages.Length > effectiveOptions.MaxPages) {
                throw new PdfReadLimitException(PdfReadLimitKind.RenderPages, effectiveOptions.MaxPages, pages.Length, "PDF render page count exceeded the configured limit.");
            }

            var results = new List<PdfPageRenderResult>(pages.Length);
            long totalOutputBytes = 0;
            for (int i = 0; i < pages.Length; i++) {
                execution.Token.ThrowIfCancellationRequested();
                PdfPageRenderResult result = RenderPage(document, pages[i], effectiveOptions, execution.Token, pixelsOnly);
                totalOutputBytes = checked(totalOutputBytes + result.OutputByteLength);
                if (totalOutputBytes > effectiveOptions.MaxTotalOutputBytes) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.RenderBytes, effectiveOptions.MaxTotalOutputBytes, totalOutputBytes);
                }
                results.Add(result);
            }

            execution.ThrowIfCancellationRequested();
            return results.AsReadOnly();
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    /// <summary>Renders parsed page ranges such as <c>1-3,5</c>.</summary>
    public static IReadOnlyList<PdfPageRenderResult> RenderPages(
        byte[] pdf,
        string pageRanges,
        PdfPageRenderOptions? options = null,
        PdfLoadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        return RenderPages(token => PdfReadDocument.Open(pdf, readOptions, token), pageRanges, options, cancellationToken);
    }

    /// <summary>Renders pages resolved by a document-relative selector.</summary>
    public static IReadOnlyList<PdfPageRenderResult> RenderPages(
        byte[] pdf,
        OfficePageSelector selector,
        PdfPageRenderOptions? options = null,
        PdfLoadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        return RenderPages(token => PdfReadDocument.Open(pdf, readOptions, token), selector, options, cancellationToken);
    }

    internal static IReadOnlyList<PdfPageRenderResult> RenderPages(
        Func<CancellationToken, PdfReadDocument> getDocument,
        PdfPageSelection? selection,
        PdfPageRenderOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(getDocument, nameof(getDocument));
        return RenderPagesCore(
            getDocument,
            pageCount => selection?.ToPageNumbers(pageCount, nameof(selection)) ?? Enumerable.Range(1, pageCount).ToArray(),
            options,
            cancellationToken);
    }

    internal static IReadOnlyList<PdfPageRenderResult> RenderPages(
        Func<CancellationToken, PdfReadDocument> getDocument,
        string pageRanges,
        PdfPageRenderOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(getDocument, nameof(getDocument));
        Guard.NotNull(pageRanges, nameof(pageRanges));
        return RenderPagesCore(
            getDocument,
            pageCount => OfficePageSelector.Parse(pageRanges).ResolveSelection(pageCount).ToPageNumbers(pageCount, nameof(pageRanges)),
            options,
            cancellationToken);
    }

    internal static IReadOnlyList<PdfPageRenderResult> RenderPages(
        Func<CancellationToken, PdfReadDocument> getDocument,
        OfficePageSelector selector,
        PdfPageRenderOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(getDocument, nameof(getDocument));
        Guard.NotNull(selector, nameof(selector));
        return RenderPagesCore(
            getDocument,
            pageCount => selector.ResolveSelection(pageCount).ToPageNumbers(pageCount, nameof(selector)),
            options,
            cancellationToken);
    }

    // Raster output exposes only pixels, which viewing permits without content-copy permission. SVG carries text and source
    // images, and caller codecs, shaping providers, and provider-loaded font programs receive source content, so those
    // still require extraction permission.
    private static bool ProducesPixelsOnly(PdfPageRenderOptions options) =>
        options.Format == global::OfficeIMO.Drawing.OfficeImageExportFormat.Png &&
        options.ImageCodec is null &&
        options.TextShapingProvider is null &&
        options.Fonts.Faces.All(face => face.HasBuiltInProgram);

    // printing paints only annotations with the Print flag, as a printed page shows them.
    internal static PdfPageRenderResult RenderPage(PdfReadDocument document, int pageNumber, PdfPageRenderOptions options, CancellationToken cancellationToken, bool forDisplay = false,
        bool printing = false) {
        var timer = Stopwatch.StartNew();
        IReadOnlyList<PdfRenderCapabilityDiagnostic> capabilityDiagnostics = Array.Empty<PdfRenderCapabilityDiagnostic>();
        try {
            cancellationToken.ThrowIfCancellationRequested();
            capabilityDiagnostics = document.Pages[pageNumber - 1].GetRenderCapabilityDiagnostics(
                options.MaxDiagnosticsPerPage, options.MaxDiagnosticCharactersPerPage, cancellationToken);
            void ConfigureDrawing(OfficeDrawing scene) {
                scene.Fonts.AddRangePreservingExisting(options.Fonts);
                scene.TextShapingProvider = options.TextShapingProvider;
                scene.TextShapingLanguage = options.TextShapingLanguage;
            }
            var pageSize = document.Pages[pageNumber - 1].GetVisualPageSize();
            double functionScale = options.GetScale(pageSize.Width, pageSize.Height);
            OfficeDrawing drawing = forDisplay ? document.Pages[pageNumber - 1].ToDisplayDrawing(cancellationToken, ConfigureDrawing, functionScale, options.MaxPixelsPerPage, printing: printing)
                : document.Pages[pageNumber - 1].ToDrawing(cancellationToken, ConfigureDrawing, functionScale, options.MaxPixelsPerPage);
            double scale = options.GetScale(drawing);
            int width = checked((int)Math.Ceiling(drawing.Width * scale));
            int height = checked((int)Math.Ceiling(drawing.Height * scale));
            long pixels = checked((long)width * height);
            if (pixels > options.MaxPixelsPerPage) {
                throw new PdfReadLimitException(PdfReadLimitKind.RenderPixels, options.MaxPixelsPerPage, pixels, "PDF render pixel count exceeded the configured per-page limit.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes;
            if (options.Format == global::OfficeIMO.Drawing.OfficeImageExportFormat.Png) {
                bytes = RenderDrawingAsPng(
                    drawing,
                    scale,
                    options.Background,
                    options.ImageCodec,
                    options.MaxPixelsPerPage,
                    options.TextShapingProvider,
                    options.TextShapingLanguage,
                    cancellationToken);
            } else {
                try {
                    bytes = OfficeDrawingSvgExporter.ToSvgBytes(
                        drawing,
                        scale,
                        OfficeSvgSizeUnit.Point,
                        imageCodec: null,
                        resourceIdPrefix: null,
                        maximumUtf8Bytes: options.MaxOutputBytesPerPage,
                        cancellationToken: cancellationToken);
                } catch (OfficeImageExportBatchLimitException exception) {
                    throw PdfReadLimitException.Create(
                        PdfReadLimitKind.RenderBytes,
                        options.MaxOutputBytesPerPage,
                        exception.Actual);
                }
            }
            if (bytes.LongLength > options.MaxOutputBytesPerPage) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.RenderBytes, options.MaxOutputBytesPerPage, bytes.LongLength);
            }
            timer.Stop();
            return new PdfPageRenderResult(pageNumber, options.Format, bytes, width, height, timer.Elapsed,
                capabilityDiagnostics, options.MaxDiagnosticCharactersPerPage);
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception ex) when (options.ContinueOnError && ex is not OutOfMemoryException && ex is not StackOverflowException) {
            timer.Stop();
            string prefix = ex.GetType().Name + ": ";
            string message = ex.Message;
            int length = (int)Math.Min(options.MaxDiagnosticCharactersPerPage,
                (long)prefix.Length + message.Length);
            var characters = new char[length];
            int prefixLength = Math.Min(prefix.Length, length);
            prefix.CopyTo(0, characters, 0, prefixLength);
            if (length > prefixLength)
                message.CopyTo(0, characters, prefixLength, length - prefixLength);
            string error = new string(characters);
            return new PdfPageRenderResult(pageNumber, options.Format, null, 0, 0, timer.Elapsed,
                capabilityDiagnostics, options.MaxDiagnosticCharactersPerPage, new[] { error });
        }
    }
}

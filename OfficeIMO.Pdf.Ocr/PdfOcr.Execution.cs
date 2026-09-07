using System.Threading;
using System.Threading.Tasks;
using System.Runtime.ExceptionServices;
using OfficeIMO.Ocr;

namespace OfficeIMO.Pdf.Ocr;

internal static partial class PdfOcr {
    // Parsing and rendering remain on one producer: parsed-page caches are not concurrent collections.
    // Only provider requests and projection of their independent results overlap.
    private static async Task<IReadOnlyList<PdfOcrPageMergeResult>> RecognizePagesAsync(
        PdfReadDocument document,
        PdfReadDocument overlapDocument,
        PdfDocumentReadResult logical,
        int[] selectedPages,
        OcrEngineExecution engine,
        PdfOcrMergeOptions options,
        CancellationToken cancellationToken) {
        int concurrency = engine.Capabilities.SupportsConcurrentRequests ? options.MaxConcurrentPages : 1;
        var renderOptions = new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Png,
            Dpi = options.Dpi,
            ImageCodec = options.ImageCodec,
            MaxPixelsPerPage = options.MaxPixelsPerPage,
            MaxOutputBytesPerPage = options.MaxRenderedBytesPerPage,
            ContinueOnError = false
        };
        renderOptions.Validate();
        var results = new PdfOcrPageMergeResult[selectedPages.Length];
        var pending = new List<Task>();
        ExceptionDispatchInfo? providerFailure = null;
        var nativePages = logical.Pages.GroupBy(static page => page.PageNumber)
            .ToDictionary(static group => group.Key, static group => group.First());
        using var workCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        try {
            for (int index = 0; index < selectedPages.Length; index++) {
                if (pending.Count >= concurrency) {
                    Task completed = await Task.WhenAny(pending).ConfigureAwait(false);
                    await completed.ConfigureAwait(false);
                    pending.Remove(completed);
                }
                workCancellation.Token.ThrowIfCancellationRequested();
                int pageNumber = selectedPages[index];
                PdfPageRenderResult render = PdfPageImageRenderer.RenderPage(
                    document, pageNumber, renderOptions, workCancellation.Token);
                if (render.Diagnostics.Count > options.MaxDiagnosticsPerPage)
                    throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxDiagnosticsPerPage, render.Diagnostics.Count);
                EnsureCharacters(render.Diagnostics, options.MaxDiagnosticCharactersPerPage);
                IReadOnlyList<PdfSelectionQuad> nativeBounds = PdfPageInteractionMap.GetOcrOverlapTextSpanBounds(
                    overlapDocument.Pages[pageNumber - 1]);
                (double width, double height) = document.Pages[pageNumber - 1].GetInteractionPageSize();
                string candidateId = "pdf-page-" + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
                var request = new OcrRequest {
                    Payload = render.Bytes!,
                    MediaType = "image/png",
                    FileName = candidateId + ".png",
                    SourceId = options.SourceId,
                    SourceName = options.SourceName,
                    CandidateId = candidateId,
                    CandidateKind = "page",
                    PageNumber = pageNumber,
                    PixelWidth = render.Width,
                    PixelHeight = render.Height,
                    Region = new OcrRegion { X = 0D, Y = 0D, Width = width, Height = height },
                    RegionCoordinateUnit = OcrCoordinateUnit.Points,
                    Language = options.Language,
                    ProviderOptions = options.ProviderOptions
                };
                pending.Add(RecognizePageAsync(index, request, render.Diagnostics, nativePages[pageNumber], nativeBounds));
            }
            await Task.WhenAll(pending).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return Array.AsReadOnly(results);
        } catch {
            workCancellation.Cancel();
            // Observe every started operation before disposing its cancellation scope.
            try { await Task.WhenAll(pending).ConfigureAwait(false); } catch { }
            if (!cancellationToken.IsCancellationRequested) providerFailure?.Throw();
            throw;
        }

        async Task RecognizePageAsync(int index, OcrRequest request, IReadOnlyList<string> renderDiagnostics,
            PdfLogicalPage nativePage, IReadOnlyList<PdfSelectionQuad> nativeBounds) {
            try {
                OcrResult recognized = await engine.RecognizeAsync(
                    request, options.ProviderTimeout, workCancellation.Token).ConfigureAwait(false);
                ProjectedOcrResult projected = ProjectResult(recognized, request, engine.Id, options, workCancellation.Token);
                if (renderDiagnostics.Count > 0) {
                    var diagnostics = new List<string>(renderDiagnostics);
                    diagnostics.AddRange(projected.Diagnostics);
                    if (diagnostics.Count > options.MaxDiagnosticsPerPage) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxDiagnosticsPerPage, diagnostics.Count);
                    }
                    EnsureCharacters(diagnostics, options.MaxDiagnosticCharactersPerPage);
                    projected = new ProjectedOcrResult(projected.Words, diagnostics.AsReadOnly(),
                        projected.Provider, projected.Model, projected.Language);
                }
                results[index] = MergePage(nativePage, nativeBounds, projected, options, workCancellation.Token);
            } catch (Exception exception) {
                if (exception is not OperationCanceledException)
                    Interlocked.CompareExchange(ref providerFailure, ExceptionDispatchInfo.Capture(exception), null);
                workCancellation.Cancel();
                throw;
            }
        }
    }
}

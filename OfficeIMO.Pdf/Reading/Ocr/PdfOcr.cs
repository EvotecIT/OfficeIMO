using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

/// <summary>Engine-owned OCR rendering and merge orchestration over an external provider.</summary>
internal static class PdfOcr {
    /// <summary>Renders selected pages, invokes the provider, and merges normalized OCR words with native text evidence.</summary>
    public static async Task<PdfOcrMergeResult> RecognizeAndMergeAsync(
        byte[] pdf,
        IPdfOcrProvider provider,
        PdfOcrMergeOptions? options = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(provider, nameof(provider));
        PdfOcrMergeOptions effectiveOptions = options ?? new PdfOcrMergeOptions();
        effectiveOptions.Validate();
        PdfReadDocument readDocument = PdfReadDocument.Open(pdf, readOptions);
        PdfLogicalDocument logical = PdfLogicalDocument.From(readDocument);
        var renderOptions = new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Png,
            Dpi = effectiveOptions.Dpi,
            MaxPages = effectiveOptions.MaxPages,
            MaxPixelsPerPage = effectiveOptions.MaxPixelsPerPage,
            ContinueOnError = false
        };
        IReadOnlyList<PdfPageRenderResult> rendered = PdfPageImageRenderer.RenderPages(pdf, effectiveOptions.Selection, renderOptions, readOptions, cancellationToken);
        var pages = new List<PdfOcrPageMergeResult>(rendered.Count);
        for (int i = 0; i < rendered.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfPageRenderResult render = rendered[i];
            PdfLogicalPage nativePage = logical.Pages.First(page => page.PageNumber == render.PageNumber);
            double scale = effectiveOptions.Dpi / 72D;
            var request = new PdfOcrRequest(render.PageNumber, render.Bytes!, render.Width, render.Height, nativePage.Width, nativePage.Height, scale);
            PdfOcrResponse response = await provider.RecognizeAsync(request, cancellationToken).ConfigureAwait(false)
                ?? throw new InvalidOperationException("OCR provider returned a null response.");
            pages.Add(MergePage(nativePage, response, request, effectiveOptions));
        }

        return new PdfOcrMergeResult(logical, pages.AsReadOnly());
    }

    private static PdfOcrPageMergeResult MergePage(PdfLogicalPage nativePage, PdfOcrResponse response, PdfOcrRequest request, PdfOcrMergeOptions options) {
        var diagnostics = new List<string>(response.Diagnostics);
        var accepted = new List<PdfRecognizedWord>();
        int lowConfidence = 0;
        int nativeOverlap = 0;
        for (int i = 0; i < response.Words.Count; i++) {
            PdfOcrWord word = response.Words[i];
            if (!IsValid(word, request)) {
                diagnostics.Add("InvalidWordGeometry: provider word " + i + " was outside the rendered page or contained non-finite values.");
                continue;
            }

            if (word.Confidence < options.MinimumConfidence) {
                lowConfidence++;
                continue;
            }

            var normalized = new PdfRecognizedWord(word.Text, word.X / request.Scale, word.Y / request.Scale, word.Width / request.Scale, word.Height / request.Scale, word.Confidence);
            if (OverlapsNativeText(normalized, nativePage.TextBlocks, nativePage.Height, options.NativeTextOverlapThreshold)) {
                nativeOverlap++;
                continue;
            }

            accepted.Add(normalized);
        }

        accepted.Sort(static (left, right) => {
            int y = left.Y.CompareTo(right.Y);
            return y != 0 ? y : left.X.CompareTo(right.X);
        });
        string text = BuildMergedText(nativePage, accepted);
        return new PdfOcrPageMergeResult(nativePage.PageNumber, accepted.AsReadOnly(), lowConfidence, nativeOverlap, diagnostics.AsReadOnly(), text);
    }

    private static bool IsValid(PdfOcrWord word, PdfOcrRequest request) =>
        IsFinite(word.X) && IsFinite(word.Y) && IsFinite(word.Width) && IsFinite(word.Height) && IsFinite(word.Confidence) &&
        word.X >= 0D && word.Y >= 0D && word.Width > 0D && word.Height > 0D && word.Confidence >= 0D && word.Confidence <= 1D &&
        word.X + word.Width <= request.PixelWidth + 0.01D && word.Y + word.Height <= request.PixelHeight + 0.01D;

    private static bool OverlapsNativeText(PdfRecognizedWord word, IReadOnlyList<PdfLogicalTextBlock> blocks, double pageHeight, double threshold) {
        double wordArea = word.Width * word.Height;
        for (int i = 0; i < blocks.Count; i++) {
            PdfLogicalTextBlock block = blocks[i];
            double blockHeight = Math.Max(block.FontSize * 1.2D, 1D);
            double blockTop = pageHeight - block.BaselineY - blockHeight;
            double overlapWidth = Math.Max(0D, Math.Min(word.X + word.Width, block.XEnd) - Math.Max(word.X, block.XStart));
            double overlapHeight = Math.Max(0D, Math.Min(word.Y + word.Height, blockTop + blockHeight) - Math.Max(word.Y, blockTop));
            if ((overlapWidth * overlapHeight) / wordArea >= threshold) return true;
        }

        return false;
    }

    private static string BuildMergedText(PdfLogicalPage page, List<PdfRecognizedWord> words) {
        var items = new List<(double Y, double X, string Text)>(page.TextBlocks.Count + words.Count);
        for (int i = 0; i < page.TextBlocks.Count; i++) {
            PdfLogicalTextBlock block = page.TextBlocks[i];
            items.Add((page.Height - block.BaselineY - block.FontSize, block.XStart, block.Text));
        }

        for (int i = 0; i < words.Count; i++) items.Add((words[i].Y, words[i].X, words[i].Text));
        return string.Join(Environment.NewLine, items.OrderBy(static item => item.Y).ThenBy(static item => item.X).Select(static item => item.Text));
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}

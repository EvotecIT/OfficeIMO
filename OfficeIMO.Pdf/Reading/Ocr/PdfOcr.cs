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
        PdfLoadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(provider, nameof(provider));
        PdfOcrMergeOptions effectiveOptions = options ?? new PdfOcrMergeOptions();
        effectiveOptions.Validate();
        PdfReadDocument readDocument = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        PdfReadDocument overlapReadDocument = readOptions?.IncludeArtifactText == true
            ? readDocument
            : PdfReadDocument.Open(pdf, PdfLoadOptions.WithArtifactText(readOptions), cancellationToken);
        int[] selectedPages = effectiveOptions.Selection?.ToPageNumbers(
            readDocument.Pages.Count,
            nameof(effectiveOptions.Selection)) ?? Enumerable.Range(1, readDocument.Pages.Count).ToArray();
        if (selectedPages.Length > effectiveOptions.MaxPages) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.Pages, effectiveOptions.MaxPages, selectedPages.Length);
        }
        PdfDocumentReadResult logical = PdfDocumentReadEngine.Read(
            readDocument,
            new PdfReadOptions {
                Profile = PdfReadProfile.Structured,
                PageSelection = effectiveOptions.Selection,
                Pipeline = CreateUnderstandingPipelineOptions(effectiveOptions)
            },
            cancellationToken);
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
            PdfReadPage readPage = readDocument.Pages[render.PageNumber - 1];
            PdfReadPage overlapReadPage = overlapReadDocument.Pages[render.PageNumber - 1];
            IReadOnlyList<PdfSelectionQuad> nativeTextBounds = PdfPageInteractionMap.GetOcrOverlapTextSpanBounds(overlapReadPage);
            (double visualWidth, double visualHeight) = readPage.GetInteractionPageSize();
            double scale = effectiveOptions.Dpi / 72D;
            var request = new PdfOcrRequest(render.PageNumber, render.Bytes!, render.Width, render.Height, visualWidth, visualHeight, scale);
            PdfOcrResponse response = await provider.RecognizeAsync(request, cancellationToken).ConfigureAwait(false)
                ?? throw new InvalidOperationException("OCR provider returned a null response.");
            pages.Add(MergePage(nativePage, nativeTextBounds, readPage, response, request, effectiveOptions, cancellationToken));
        }

        IReadOnlyList<PdfOcrPageMergeResult> mergedPages = pages.AsReadOnly();
        PdfDocumentReadResult enriched = PdfOcrLogicalDocumentBuilder.Build(
            logical,
            mergedPages,
            effectiveOptions,
            cancellationToken);
        return new PdfOcrMergeResult(logical, enriched, mergedPages);
    }

    internal static PdfUnderstandingPipelineOptions CreateUnderstandingPipelineOptions(PdfOcrMergeOptions options) {
        Guard.NotNull(options, nameof(options));
        return new PdfUnderstandingPipelineOptions { MaxPages = options.MaxPages };
    }

    private static PdfOcrPageMergeResult MergePage(PdfLogicalPage nativePage, IReadOnlyList<PdfSelectionQuad> nativeTextBounds, PdfReadPage readPage, PdfOcrResponse response, PdfOcrRequest request, PdfOcrMergeOptions options, CancellationToken cancellationToken) {
        ValidateProviderResponse(nativePage, response, options);
        var diagnostics = new List<string>(response.Diagnostics);
        var accepted = new List<PdfRecognizedWord>();
        int lowConfidence = 0;
        int nativeOverlap = 0;
        long overlapComparisons = 0;
        for (int i = 0; i < response.Words.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfOcrWord word = response.Words[i];
            if (!IsValid(word, request)) {
                diagnostics.Add("InvalidWordGeometry: provider word " + i + " was outside the rendered page or contained non-finite values.");
                continue;
            }

            if (word.Confidence < options.MinimumConfidence) {
                lowConfidence++;
                continue;
            }

            var normalized = new PdfRecognizedWord(word.Text, word.X / request.Scale, word.Y / request.Scale, word.Width / request.Scale, word.Height / request.Scale, word.Confidence, i);
            if (OverlapsNativeText(
                    normalized,
                    nativeTextBounds,
                    options.NativeTextOverlapThreshold,
                    options.MaxNativeTextOverlapComparisonsPerPage,
                    ref overlapComparisons,
                    cancellationToken)) {
                nativeOverlap++;
                continue;
            }

            accepted.Add(normalized);
        }

        accepted.Sort(static (left, right) => {
            int y = left.Y.CompareTo(right.Y);
            return y != 0 ? y : left.X.CompareTo(right.X);
        });
        string text = BuildMergedText(nativePage, readPage, accepted, options.MaxMergedTextCharactersPerPage, cancellationToken);
        return new PdfOcrPageMergeResult(nativePage.PageNumber, accepted.AsReadOnly(), lowConfidence, nativeOverlap, diagnostics.AsReadOnly(), text, response.Provider, response.Model, response.Language);
    }

    private static bool IsValid(PdfOcrWord word, PdfOcrRequest request) =>
        IsFinite(word.X) && IsFinite(word.Y) && IsFinite(word.Width) && IsFinite(word.Height) && IsFinite(word.Confidence) &&
        word.X >= 0D && word.Y >= 0D && word.Width > 0D && word.Height > 0D && word.Confidence >= 0D && word.Confidence <= 1D &&
        word.X + word.Width <= request.PixelWidth + 0.01D && word.Y + word.Height <= request.PixelHeight + 0.01D;

    private static bool OverlapsNativeText(
        PdfRecognizedWord word,
        IReadOnlyList<PdfSelectionQuad> nativeTextBounds,
        double threshold,
        long maximumComparisons,
        ref long comparisons,
        CancellationToken cancellationToken) {
        double wordArea = word.Width * word.Height;
        double requiredArea = wordArea * threshold;
        var intersections = new List<OcrOverlapRectangle>();
        double summedIntersectionArea = 0D;
        for (int i = 0; i < nativeTextBounds.Count; i++) {
            comparisons = checked(comparisons + 1L);
            if (comparisons > maximumComparisons) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, maximumComparisons, comparisons);
            }
            if ((i & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            PdfSelectionQuad bounds = nativeTextBounds[i];
            double left = Math.Max(word.X, bounds.Left);
            double top = Math.Max(word.Y, bounds.Top);
            double right = Math.Min(word.X + word.Width, bounds.Right);
            double bottom = Math.Min(word.Y + word.Height, bounds.Bottom);
            double overlapWidth = Math.Max(0D, right - left);
            double overlapHeight = Math.Max(0D, bottom - top);
            double overlapArea = overlapWidth * overlapHeight;
            if (overlapArea >= requiredArea) return true;
            if (overlapArea <= 0D) continue;
            intersections.Add(new OcrOverlapRectangle(left, top, right, bottom));
            summedIntersectionArea += overlapArea;
        }

        return summedIntersectionArea >= requiredArea &&
            CalculateRectangleUnionArea(intersections, cancellationToken) >= requiredArea;
    }

    private static double CalculateRectangleUnionArea(
        IReadOnlyList<OcrOverlapRectangle> rectangles,
        CancellationToken cancellationToken) {
        if (rectangles.Count == 0) return 0D;
        var yCoordinates = new List<double>(checked(rectangles.Count * 2));
        for (int i = 0; i < rectangles.Count; i++) {
            yCoordinates.Add(rectangles[i].Top);
            yCoordinates.Add(rectangles[i].Bottom);
        }
        yCoordinates.Sort();
        int uniqueCount = 0;
        for (int i = 0; i < yCoordinates.Count; i++) {
            if (uniqueCount == 0 || yCoordinates[i] != yCoordinates[uniqueCount - 1]) {
                yCoordinates[uniqueCount++] = yCoordinates[i];
            }
        }
        if (uniqueCount < yCoordinates.Count) yCoordinates.RemoveRange(uniqueCount, yCoordinates.Count - uniqueCount);

        var coordinateIndexes = new Dictionary<double, int>(yCoordinates.Count);
        for (int i = 0; i < yCoordinates.Count; i++) coordinateIndexes.Add(yCoordinates[i], i);
        var events = new List<OcrOverlapEvent>(checked(rectangles.Count * 2));
        for (int i = 0; i < rectangles.Count; i++) {
            OcrOverlapRectangle rectangle = rectangles[i];
            int topIndex = coordinateIndexes[rectangle.Top];
            int bottomIndex = coordinateIndexes[rectangle.Bottom] - 1;
            events.Add(new OcrOverlapEvent(rectangle.Left, topIndex, bottomIndex, 1));
            events.Add(new OcrOverlapEvent(rectangle.Right, topIndex, bottomIndex, -1));
        }
        events.Sort(static (first, second) => first.X.CompareTo(second.X));

        var coverage = new OcrVerticalCoverageTree(yCoordinates);
        double area = 0D;
        double previousX = events[0].X;
        int eventIndex = 0;
        while (eventIndex < events.Count) {
            if ((eventIndex & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            double x = events[eventIndex].X;
            area += (x - previousX) * coverage.CoveredLength;
            while (eventIndex < events.Count && events[eventIndex].X == x) {
                OcrOverlapEvent current = events[eventIndex++];
                coverage.Update(current.TopIndex, current.BottomIndex, current.Delta);
            }
            previousX = x;
        }
        return area;
    }

    private readonly struct OcrOverlapEvent {
        internal OcrOverlapEvent(double x, int topIndex, int bottomIndex, int delta) {
            X = x;
            TopIndex = topIndex;
            BottomIndex = bottomIndex;
            Delta = delta;
        }

        internal double X { get; }
        internal int TopIndex { get; }
        internal int BottomIndex { get; }
        internal int Delta { get; }
    }

    private sealed class OcrVerticalCoverageTree {
        private readonly IReadOnlyList<double> _coordinates;
        private readonly int[] _coverageCounts;
        private readonly double[] _coveredLengths;

        internal OcrVerticalCoverageTree(IReadOnlyList<double> coordinates) {
            _coordinates = coordinates;
            int intervalCount = coordinates.Count - 1;
            int storageSize = checked(Math.Max(1, intervalCount) * 4);
            _coverageCounts = new int[storageSize];
            _coveredLengths = new double[storageSize];
        }

        internal double CoveredLength => _coveredLengths[1];

        internal void Update(int firstInterval, int lastInterval, int delta) {
            if (firstInterval > lastInterval) return;
            Update(1, 0, _coordinates.Count - 2, firstInterval, lastInterval, delta);
        }

        private void Update(int node, int left, int right, int firstInterval, int lastInterval, int delta) {
            if (firstInterval <= left && right <= lastInterval) {
                _coverageCounts[node] += delta;
            } else {
                int middle = left + ((right - left) / 2);
                if (firstInterval <= middle) Update(node * 2, left, middle, firstInterval, lastInterval, delta);
                if (lastInterval > middle) Update((node * 2) + 1, middle + 1, right, firstInterval, lastInterval, delta);
            }

            if (_coverageCounts[node] > 0) {
                _coveredLengths[node] = _coordinates[right + 1] - _coordinates[left];
            } else if (left == right) {
                _coveredLengths[node] = 0D;
            } else {
                _coveredLengths[node] = _coveredLengths[node * 2] + _coveredLengths[(node * 2) + 1];
            }
        }
    }

    private readonly struct OcrOverlapRectangle {
        internal OcrOverlapRectangle(double left, double top, double right, double bottom) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        internal double Left { get; }
        internal double Top { get; }
        internal double Right { get; }
        internal double Bottom { get; }
    }

    private static string BuildMergedText(
        PdfLogicalPage page,
        PdfReadPage readPage,
        List<PdfRecognizedWord> words,
        int maximumCharacters,
        CancellationToken cancellationToken) {
        IReadOnlyList<PdfOcrLogicalTextLine> ocrLines = PdfOcrLogicalDocumentBuilder.BuildTextLines(words, cancellationToken);
        var items = new List<(double Y, double X, string Text)>(page.TextBlocks.Count + ocrLines.Count);
        for (int i = 0; i < page.TextBlocks.Count; i++) {
            PdfLogicalTextBlock block = page.TextBlocks[i];
            double blockHeight = Math.Max(block.FontSize, 1D);
            PdfVisualBounds bounds = readPage.TransformBoundsToVisual(
                Math.Min(block.XStart, block.XEnd),
                block.BaselineY,
                Math.Max(block.XStart, block.XEnd),
                block.BaselineY + blockHeight);
            items.Add((bounds.Top, bounds.Left, block.Text));
        }

        for (int i = 0; i < ocrLines.Count; i++) items.Add((ocrLines[i].Top, ocrLines[i].Left, ocrLines[i].Text));
        var builder = new System.Text.StringBuilder(Math.Min(maximumCharacters, 4096));
        foreach ((double _, double _, string text) in items.OrderBy(static item => item.Y).ThenBy(static item => item.X)) {
            int separatorLength = builder.Length == 0 ? 0 : Environment.NewLine.Length;
            long projected = (long)builder.Length + separatorLength + text.Length;
            if (projected > maximumCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, maximumCharacters, projected);
            }
            if (separatorLength > 0) builder.AppendLine();
            builder.Append(text);
        }
        return builder.ToString();
    }

    private static void ValidateProviderResponse(PdfLogicalPage nativePage, PdfOcrResponse response, PdfOcrMergeOptions options) {
        if (response.Words.Count > options.MaxOcrWordsPerPage) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxOcrWordsPerPage, response.Words.Count);
        }
        if (response.Diagnostics.Count > options.MaxDiagnosticsPerPage) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxDiagnosticsPerPage, response.Diagnostics.Count);
        }
        if (nativePage.TextBlocks.Count > options.MaxNativeTextBlocksPerPage) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxNativeTextBlocksPerPage, nativePage.TextBlocks.Count);
        }
        EnsureCharacters(response.Words.Select(static word => word.Text), options.MaxOcrTextCharactersPerPage);
        EnsureCharacters(response.Diagnostics, options.MaxDiagnosticCharactersPerPage);
    }

    private static void EnsureCharacters(IEnumerable<string> values, int maximum) {
        long total = 0;
        foreach (string value in values) {
            total = checked(total + value.Length);
            if (total > maximum) throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, maximum, total);
        }
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}

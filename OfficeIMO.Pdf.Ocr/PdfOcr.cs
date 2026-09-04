using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>PDF-specific rendering and native-text merge orchestration over an engine-neutral OCR provider.</summary>
internal static class PdfOcr {
    internal static async Task<PdfOcrMergeResult> RecognizeAndMergeAsync(
        byte[] pdf,
        IOcrEngine engine,
        PdfOcrMergeOptions? options = null,
        PdfLoadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(engine, nameof(engine));
        OcrEngineExecution engineExecution = OcrEngineRunner.CreateExecution(engine);
        string engineId = engineExecution.Id;
        EnsurePngSupport(engineExecution.Capabilities);
        PdfOcrMergeOptions effectiveOptions = options?.Clone() ?? new PdfOcrMergeOptions();
        effectiveOptions.Validate();
        PdfReadOptions semanticOptions = effectiveOptions.ReadOptions;
        PdfReadDocument readDocument = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        PdfReadDocument overlapReadDocument = readOptions?.IncludeArtifactText == true
            ? readDocument
            : PdfReadDocument.Open(pdf, PdfLoadOptions.WithArtifactText(readOptions), cancellationToken);
        int[] selectedPages = semanticOptions.PageSelection?.ToPageNumbers(
            readDocument.Pages.Count,
            nameof(semanticOptions.PageSelection)) ?? Enumerable.Range(1, readDocument.Pages.Count).ToArray();
        if (selectedPages.Length > effectiveOptions.MaxPages) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.Pages, effectiveOptions.MaxPages, selectedPages.Length);
        }

        PdfTextLayoutOptions layoutOptions = semanticOptions.LayoutOptions;
        PdfUnderstandingPipelineOptions pipelineOptions = PdfUnderstandingPipelineOptions.Resolve(semanticOptions.Pipeline);
        PdfDocumentReadResult logical = PdfDocumentReadEngine.Read(
            readDocument,
            semanticOptions,
            out IReadOnlyList<PdfUnderstandingPageResult> pageAnalyses,
            cancellationToken);
        var renderOptions = new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Png,
            Dpi = effectiveOptions.Dpi,
            MaxPages = effectiveOptions.MaxPages,
            MaxPixelsPerPage = effectiveOptions.MaxPixelsPerPage,
            ContinueOnError = false
        };
        IReadOnlyList<PdfPageRenderResult> rendered = PdfPageImageRenderer.RenderPages(
            pdf,
            semanticOptions.PageSelection,
            renderOptions,
            readOptions,
            cancellationToken);
        var pages = new List<PdfOcrPageMergeResult>(rendered.Count);
        for (int index = 0; index < rendered.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfPageRenderResult render = rendered[index];
            PdfLogicalPage nativePage = logical.Pages.First(page => page.PageNumber == render.PageNumber);
            PdfReadPage readPage = readDocument.Pages[render.PageNumber - 1];
            PdfReadPage overlapReadPage = overlapReadDocument.Pages[render.PageNumber - 1];
            IReadOnlyList<PdfSelectionQuad> nativeTextBounds = PdfPageInteractionMap.GetOcrOverlapTextSpanBounds(overlapReadPage);
            (double visualWidth, double visualHeight) = readPage.GetInteractionPageSize();
            byte[] payload = (byte[])render.Bytes!.Clone();
            string candidateId = "pdf-page-" + render.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var request = new OcrRequest {
                Payload = payload,
                MediaType = "image/png",
                FileName = candidateId + ".png",
                SourceId = effectiveOptions.SourceId,
                SourceName = effectiveOptions.SourceName,
                CandidateId = candidateId,
                CandidateKind = "page",
                PageNumber = render.PageNumber,
                PixelWidth = render.Width,
                PixelHeight = render.Height,
                Region = new OcrRegion { X = 0D, Y = 0D, Width = visualWidth, Height = visualHeight },
                RegionCoordinateUnit = OcrCoordinateUnit.Points,
                Language = effectiveOptions.Language,
                ProviderOptions = effectiveOptions.ProviderOptions
            };
            OcrResult result = await engineExecution.RecognizeAsync(
                request,
                effectiveOptions.ProviderTimeout,
                cancellationToken).ConfigureAwait(false);
            ProjectedOcrResult projected = ProjectResult(result, request, engineId, effectiveOptions, cancellationToken);
            pages.Add(MergePage(nativePage, nativeTextBounds, projected, effectiveOptions, cancellationToken));
        }

        var mergedPages = pages.AsReadOnly();
        PdfDocumentReadResult enriched = PdfOcrLogicalDocumentBuilder.Build(
            readDocument,
            logical,
            pageAnalyses,
            mergedPages,
            layoutOptions,
            pipelineOptions,
            cancellationToken);
        var canonicalTextByPage = new Dictionary<int, Queue<string>>();
        for (int pageIndex = 0; pageIndex < enriched.Pages.Count; pageIndex++) {
            PdfLogicalPage page = enriched.Pages[pageIndex];
            if (!canonicalTextByPage.TryGetValue(page.PageNumber, out Queue<string>? texts)) {
                texts = new Queue<string>();
                canonicalTextByPage.Add(page.PageNumber, texts);
            }
            texts.Enqueue(PdfDocumentReadResult.GetCanonicalPageText(page));
        }
        var canonicalPages = new PdfOcrPageMergeResult[mergedPages.Count];
        for (int pageIndex = 0; pageIndex < mergedPages.Count; pageIndex++) {
            PdfOcrPageMergeResult page = mergedPages[pageIndex];
            string text = canonicalTextByPage.TryGetValue(page.PageNumber, out Queue<string>? texts) && texts.Count > 0
                ? texts.Dequeue()
                : string.Empty;
            canonicalPages[pageIndex] = page.WithCanonicalText(text, effectiveOptions.MaxMergedTextCharactersPerPage);
        }
        return new PdfOcrMergeResult(logical, enriched, Array.AsReadOnly(canonicalPages));
    }

    private static ProjectedOcrResult ProjectResult(
        OcrResult result,
        OcrRequest request,
        string engineId,
        PdfOcrMergeOptions options,
        CancellationToken cancellationToken) {
        IReadOnlyList<OcrDiagnostic> returnedDiagnostics = result.Diagnostics ?? Array.Empty<OcrDiagnostic>();
        if (returnedDiagnostics.Count > options.MaxDiagnosticsPerPage) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxDiagnosticsPerPage, returnedDiagnostics.Count);
        }
        var diagnostics = new List<string>(returnedDiagnostics.Count);
        long returnedDiagnosticCharacters = 0;
        for (int diagnosticIndex = 0; diagnosticIndex < returnedDiagnostics.Count; diagnosticIndex++) {
            OcrDiagnostic diagnostic = returnedDiagnostics[diagnosticIndex];
            if (diagnostic == null) continue;
            string code = diagnostic.Code ?? string.Empty;
            string message = diagnostic.Message ?? string.Empty;
            returnedDiagnosticCharacters = AddCharacters(
                returnedDiagnosticCharacters,
                code.Length,
                options.MaxDiagnosticCharactersPerPage);
            bool hasCode = !string.IsNullOrWhiteSpace(code);
            returnedDiagnosticCharacters = AddCharacters(
                returnedDiagnosticCharacters,
                message.Length,
                options.MaxDiagnosticCharactersPerPage);
            if (hasCode) {
                returnedDiagnosticCharacters = AddCharacters(
                    returnedDiagnosticCharacters,
                    2,
                    options.MaxDiagnosticCharactersPerPage);
            }
            diagnostics.Add(hasCode ? code + ": " + message : message);
        }

        string returnedText = result.Text ?? string.Empty;
        EnsureCharacters(new[] { returnedText }, options.MaxOcrTextCharactersPerPage);
        IReadOnlyList<OcrTextSpan> returnedSpans = result.Spans ?? Array.Empty<OcrTextSpan>();
        if (returnedSpans.Count > options.MaxOcrSpansPerPage) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxOcrSpansPerPage, returnedSpans.Count);
        }
        OcrTextSpan[] spans = returnedSpans
            .Where(static span => span != null)
            .OrderBy(static span => span.Sequence)
            .ToArray();
        OcrTextSpan[] selected = spans.Where(static span => span.Level == OcrTextSpanLevel.Word).ToArray();
        if (selected.Length == 0 && options.UseLineSpansWhenWordsUnavailable) {
            selected = spans.Where(static span => span.Level == OcrTextSpanLevel.Line).ToArray();
        }
        if (selected.Length > options.MaxOcrWordsPerPage) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxOcrWordsPerPage, selected.Length);
        }

        var words = new List<ProjectedOcrWord>(selected.Length);
        long inspectedSpanCharacters = 0;
        long inspectedHierarchyCharacters = 0;
        bool usedFallbackConfidence = false;
        bool discardedHierarchyId = false;
        for (int index = 0; index < selected.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OcrTextSpan span = selected[index];
            string spanText = span.Text ?? string.Empty;
            inspectedSpanCharacters = AddCharacters(
                inspectedSpanCharacters,
                spanText.Length,
                options.MaxOcrTextCharactersPerPage);
            if (string.IsNullOrWhiteSpace(spanText)) continue;
            if (span.PageNumber.HasValue && span.PageNumber.Value != request.PageNumber && span.PageNumber.Value != 1) continue;
            if (span.Region == null || !TryConvertRegion(span.Region, span.CoordinateUnit, request, out double x, out double y, out double width, out double height)) {
                diagnostics.Add("ocr-span-geometry: A recognized span did not contain valid page geometry.");
                continue;
            }
            double confidence = span.Confidence ?? result.Confidence ?? options.ConfidenceWhenUnavailable;
            if (!IsFinite(confidence) || confidence < 0D || confidence > 1D) {
                diagnostics.Add("ocr-confidence-invalid: A recognized span reported an invalid confidence value.");
                continue;
            }
            if (!span.Confidence.HasValue && !result.Confidence.HasValue) usedFallbackConfidence = true;
            words.Add(new ProjectedOcrWord(
                spanText.Trim(),
                x,
                y,
                width,
                height,
                confidence,
                index,
                NormalizeHierarchyId(
                    span.BlockId,
                    ref inspectedHierarchyCharacters,
                    options.MaxOcrHierarchyCharactersPerPage,
                    ref discardedHierarchyId),
                NormalizeHierarchyId(
                    span.ParagraphId,
                    ref inspectedHierarchyCharacters,
                    options.MaxOcrHierarchyCharactersPerPage,
                    ref discardedHierarchyId),
                NormalizeHierarchyId(
                    span.LineId,
                    ref inspectedHierarchyCharacters,
                    options.MaxOcrHierarchyCharactersPerPage,
                    ref discardedHierarchyId)));
        }

        if (words.Count == 0 && !string.IsNullOrWhiteSpace(returnedText)) {
            diagnostics.Add("ocr-span-geometry-missing: The OCR engine returned text without usable word or line geometry, so it could not be placed on the PDF page.");
        }
        if (usedFallbackConfidence) {
            diagnostics.Add("ocr-confidence-unavailable: The OCR engine did not report confidence; the configured fallback confidence was used.");
        }
        if (discardedHierarchyId) {
            diagnostics.Add("ocr-hierarchy-id-limit: One or more OCR hierarchy identifiers exceeded 256 characters and were discarded.");
        }

        EnsureCharacters(words.Select(static word => word.Text), options.MaxOcrTextCharactersPerPage);
        EnsureCharacters(diagnostics, options.MaxDiagnosticCharactersPerPage);
        string providerValue = NormalizeProviderMetadata(result.Provider, options.MaxProviderMetadataCharactersPerPage)
            ?? engineId;
        string? modelValue = NormalizeProviderMetadata(result.Model, options.MaxProviderMetadataCharactersPerPage);
        string? languageValue = NormalizeProviderMetadata(result.Language, options.MaxProviderMetadataCharactersPerPage)
            ?? NormalizeProviderMetadata(options.Language, options.MaxProviderMetadataCharactersPerPage);
        EnsureCharacters(
            new[] { providerValue, modelValue ?? string.Empty, languageValue ?? string.Empty },
            options.MaxProviderMetadataCharactersPerPage);
        return new ProjectedOcrResult(
            words.AsReadOnly(),
            diagnostics.AsReadOnly(),
            providerValue,
            modelValue,
            languageValue);
    }

    private static PdfOcrPageMergeResult MergePage(
        PdfLogicalPage nativePage,
        IReadOnlyList<PdfSelectionQuad> nativeTextBounds,
        ProjectedOcrResult result,
        PdfOcrMergeOptions options,
        CancellationToken cancellationToken) {
        if (nativePage.TextBlocks.Count > options.MaxNativeTextBlocksPerPage) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxNativeTextBlocksPerPage, nativePage.TextBlocks.Count);
        }
        var accepted = new List<PdfRecognizedWord>(result.Words.Count);
        int lowConfidence = 0;
        int nativeOverlap = 0;
        long overlapComparisons = 0;
        for (int index = 0; index < result.Words.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            ProjectedOcrWord word = result.Words[index];
            if (word.Confidence < options.MinimumConfidence) {
                lowConfidence++;
                continue;
            }
            var normalized = new PdfRecognizedWord(
                word.Text,
                word.X,
                word.Y,
                word.Width,
                word.Height,
                word.Confidence,
                word.Sequence,
                word.BlockId,
                word.ParagraphId,
                word.LineId);
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
        return new PdfOcrPageMergeResult(
            nativePage.PageNumber,
            accepted.AsReadOnly(),
            lowConfidence,
            nativeOverlap,
            result.Diagnostics,
            string.Empty,
            result.Provider,
            result.Model,
            result.Language);
    }

    private static bool TryConvertRegion(
        OcrRegion region,
        OcrCoordinateUnit unit,
        OcrRequest request,
        out double x,
        out double y,
        out double width,
        out double height) {
        x = region.X;
        y = region.Y;
        width = region.Width;
        height = region.Height;
        if (request.Region == null || request.RegionCoordinateUnit != OcrCoordinateUnit.Points ||
            !request.PixelWidth.HasValue || !request.PixelHeight.HasValue ||
            request.Region.Width <= 0D || request.Region.Height <= 0D ||
            request.PixelWidth.Value <= 0 || request.PixelHeight.Value <= 0) return false;
        double pageWidth = request.Region.Width;
        double pageHeight = request.Region.Height;
        switch (unit) {
            case OcrCoordinateUnit.Pixels:
                double scaleX = request.PixelWidth.Value / pageWidth;
                double scaleY = request.PixelHeight.Value / pageHeight;
                x /= scaleX;
                y /= scaleY;
                width /= scaleX;
                height /= scaleY;
                break;
            case OcrCoordinateUnit.Points:
                break;
            case OcrCoordinateUnit.Normalized:
                x *= pageWidth;
                y *= pageHeight;
                width *= pageWidth;
                height *= pageHeight;
                break;
            default:
                return false;
        }
        return IsFinite(x) && IsFinite(y) && IsFinite(width) && IsFinite(height) &&
            x >= 0D && y >= 0D && width > 0D && height > 0D &&
            x + width <= pageWidth + 0.01D && y + height <= pageHeight + 0.01D;
    }

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
        for (int index = 0; index < nativeTextBounds.Count; index++) {
            comparisons = checked(comparisons + 1L);
            if (comparisons > maximumComparisons) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, maximumComparisons, comparisons);
            }
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            PdfSelectionQuad bounds = nativeTextBounds[index];
            double left = Math.Max(word.X, bounds.Left);
            double top = Math.Max(word.Y, bounds.Top);
            double right = Math.Min(word.X + word.Width, bounds.Right);
            double bottom = Math.Min(word.Y + word.Height, bounds.Bottom);
            double overlapArea = Math.Max(0D, right - left) * Math.Max(0D, bottom - top);
            if (overlapArea >= requiredArea) return true;
            if (overlapArea <= 0D) continue;
            intersections.Add(new OcrOverlapRectangle(left, top, right, bottom));
            summedIntersectionArea += overlapArea;
        }
        return summedIntersectionArea >= requiredArea && CalculateRectangleUnionArea(intersections, cancellationToken) >= requiredArea;
    }

    private static double CalculateRectangleUnionArea(IReadOnlyList<OcrOverlapRectangle> rectangles, CancellationToken cancellationToken) {
        if (rectangles.Count == 0) return 0D;
        var yCoordinates = new List<double>(checked(rectangles.Count * 2));
        for (int index = 0; index < rectangles.Count; index++) {
            yCoordinates.Add(rectangles[index].Top);
            yCoordinates.Add(rectangles[index].Bottom);
        }
        yCoordinates.Sort();
        int uniqueCount = 0;
        for (int index = 0; index < yCoordinates.Count; index++) {
            if (uniqueCount == 0 || yCoordinates[index] != yCoordinates[uniqueCount - 1]) yCoordinates[uniqueCount++] = yCoordinates[index];
        }
        if (uniqueCount < yCoordinates.Count) yCoordinates.RemoveRange(uniqueCount, yCoordinates.Count - uniqueCount);
        var coordinateIndexes = new Dictionary<double, int>(yCoordinates.Count);
        for (int index = 0; index < yCoordinates.Count; index++) coordinateIndexes.Add(yCoordinates[index], index);
        var events = new List<OcrOverlapEvent>(checked(rectangles.Count * 2));
        for (int index = 0; index < rectangles.Count; index++) {
            OcrOverlapRectangle rectangle = rectangles[index];
            events.Add(new OcrOverlapEvent(rectangle.Left, coordinateIndexes[rectangle.Top], coordinateIndexes[rectangle.Bottom] - 1, 1));
            events.Add(new OcrOverlapEvent(rectangle.Right, coordinateIndexes[rectangle.Top], coordinateIndexes[rectangle.Bottom] - 1, -1));
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

    private static string? NormalizeHierarchyId(
        string? value,
        ref long inspectedCharacters,
        int maximumCharacters,
        ref bool discarded) {
        if (string.IsNullOrEmpty(value)) return null;
        string raw = value!;
        inspectedCharacters = AddCharacters(inspectedCharacters, raw.Length, maximumCharacters);
        if (raw.Length > 256) {
            discarded = true;
            return null;
        }
        string normalized = raw.Trim();
        return normalized.Length == 0 ? null : normalized;
    }

    private static string? NormalizeProviderMetadata(string? value, int maximumCharacters) {
        if (string.IsNullOrEmpty(value)) return null;
        string raw = value!;
        if (raw.Length > maximumCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, maximumCharacters, raw.Length);
        }
        string normalized = raw.Trim();
        return normalized.Length == 0 ? null : normalized;
    }

    private static void EnsurePngSupport(OcrEngineCapabilities? capabilities) {
        IReadOnlyList<string> supported = capabilities?.SupportedMediaTypes ?? Array.Empty<string>();
        if (supported.Count == 0) return;
        if (supported.Any(mediaType =>
                string.Equals(mediaType, "image/png", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(mediaType, "image/*", StringComparison.OrdinalIgnoreCase))) return;
        throw new NotSupportedException("The OCR engine does not advertise support for rendered PNG pages.");
    }

    private static void EnsureCharacters(IEnumerable<string> values, int maximum) {
        long total = 0;
        foreach (string value in values) {
            total = AddCharacters(total, value?.Length ?? 0, maximum);
        }
    }

    private static long AddCharacters(long total, int count, int maximum) {
        long updated = checked(total + count);
        if (updated > maximum) throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, maximum, updated);
        return updated;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private sealed class ProjectedOcrResult {
        internal ProjectedOcrResult(
            IReadOnlyList<ProjectedOcrWord> words,
            IReadOnlyList<string> diagnostics,
            string provider,
            string? model,
            string? language) {
            Words = words;
            Diagnostics = diagnostics;
            Provider = provider;
            Model = model;
            Language = language;
        }
        internal IReadOnlyList<ProjectedOcrWord> Words { get; }
        internal IReadOnlyList<string> Diagnostics { get; }
        internal string Provider { get; }
        internal string? Model { get; }
        internal string? Language { get; }
    }

    private sealed class ProjectedOcrWord {
        internal ProjectedOcrWord(string text, double x, double y, double width, double height, double confidence, int sequence, string? blockId, string? paragraphId, string? lineId) {
            Text = text;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Confidence = confidence;
            Sequence = sequence;
            BlockId = blockId;
            ParagraphId = paragraphId;
            LineId = lineId;
        }
        internal string Text { get; }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal double Confidence { get; }
        internal int Sequence { get; }
        internal string? BlockId { get; }
        internal string? ParagraphId { get; }
        internal string? LineId { get; }
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
}

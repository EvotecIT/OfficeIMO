namespace OfficeIMO.Pdf;

internal static partial class PdfTextEditor {
    internal static PdfRegionText Inspect(byte[] pdf, PdfPageRegion region, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(region, nameof(region));
        PdfReadDocument document = OpenForVisualTextEditing(pdf, readOptions);
        ValidatePage(region.PageNumber, document.Pages.Count, nameof(region));
        PdfReadPage page = document.Pages[region.PageNumber - 1];
        (double originX, double originY) = page.GetPageBoundaryOrigin();
        PdfPageRegion sourceRegion = OffsetRegion(region, originX, originY);
        return OffsetRegionText(InspectSource(page, sourceRegion), -originX, -originY);
    }

    internal static IReadOnlyList<PdfTextMatch> Find(byte[] pdf, string text, PdfTextSearchOptions? options, PdfLoadOptions? readOptions) {
        return FindHits(pdf, text, options, readOptions).Select(static hit => hit.Match).ToArray();
    }

    /// <summary>Finds text page by page, reporting each page's matches, in order, as soon as the page is searched.</summary>
    internal static IReadOnlyList<PdfTextMatch> Find(byte[] pdf, string text, PdfTextSearchOptions? options, PdfLoadOptions? readOptions,
        Action<int, int, IReadOnlyList<PdfTextMatch>> pageSearched, System.Threading.CancellationToken cancellationToken) =>
        FindHits(pdf, text, options, readOptions, cancellationToken: cancellationToken,
            pageSearched: (searched, count, hits) => pageSearched(searched, count, hits.Select(static hit => hit.Match).ToArray()))
            .Select(static hit => hit.Match).ToArray();

    internal static IReadOnlyList<PdfTextMatch> Find(byte[] pdf, string text, PdfTextSearchOptions? options,
        PdfLoadOptions? readOptions, PdfRedactionSearchWorkBudget workBudget,
        System.Threading.CancellationToken cancellationToken = default) =>
        FindHits(pdf, text, options, readOptions, workBudget, cancellationToken: cancellationToken)
            .Select(static hit => hit.Match).ToArray();

    internal static TextMutationResult Add(byte[] pdf, PdfPageRegion region, string text, PdfTextEditOptions? options, PdfLoadOptions? readOptions) {
        Guard.NotNull(text, nameof(text));
        if (!HasRenderableTextLine(text)) throw new ArgumentException("Added text must contain at least one renderable line.", nameof(text));
        region = TranslateRegionToSource(pdf, region, readOptions);
        PdfTextEditOptions snapshot = (options ?? new PdfTextEditOptions()).Snapshot();
        PdfResolvedTextStyle style = ResolveStyle(snapshot, detected: null);
        style = FitStyleToRegion(style, text, region, snapshot, out string? fitWarning);
        double baselineY = region.Top - style.FontSize;
        byte[] output = StampLines(pdf, region.PageNumber, region.X, baselineY, text, style, readOptions);
        return new TextMutationResult(output, 1, fitWarning is null ? Array.Empty<string>() : new[] { fitWarning });
    }

    internal static TextMutationResult Replace(byte[] pdf, PdfPageRegion region, string text, PdfTextEditOptions? options, PdfLoadOptions? readOptions) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(region, nameof(region));
        var source = new TextEditSource(pdf, readOptions);
        region = TranslateRegionToSource(source.VisualDocument, region);
        PdfTextEditOptions snapshot = (options ?? new PdfTextEditOptions()).Snapshot();
        PdfReadPage page = source.VisualDocument.Pages[region.PageNumber - 1];
        EnsureRegionIsSafelyEditable(page, region, snapshot.AllowTextRenderingMode3);
        PdfRegionText detected = InspectSource(page, region, snapshot.AllowTextRenderingMode3);
        EnsureCompatibleRenderingModes(detected.Spans, snapshot.AllowTextRenderingMode3);
        EnsureAppendOrderIsSafe(page, detected.Spans);
        PdfResolvedTextStyle style = ResolveStyle(snapshot, detected, source.Document.Pages[region.PageNumber - 1]);
        string replacementText = PreserveAuthoredEdgeWhitespace(detected.Spans, text);
        style = FitStyleToRegion(style, replacementText, region, snapshot, out string? fitWarning);
        TextRemovalResult removal = detected.Spans.Count == 0
            ? new TextRemovalResult(pdf.ToArray(), Array.Empty<string>(), Array.Empty<PdfStamper.TextStampRequest>())
            : RemoveTextPreservingUnmatchedSpans(pdf, new[] { region.ToRedactionArea() }, readOptions,
                allowTextRenderingMode3: snapshot.AllowTextRenderingMode3, sourceDocument: source.Document);
        var requests = new List<PdfStamper.TextStampRequest>(removal.Restamps);
        var replacementRequests = new List<PdfStamper.TextStampRequest>();
        if (text.Length > 0) AddStampLines(replacementRequests, region.PageNumber, detected.Spans.Count == 0 ? region.X : detected.BaselineX, detected.Spans.Count == 0 ? region.Top - style.FontSize : detected.BaselineY, replacementText, style, detected.Spans.Count == 0 ? double.MaxValue : detected.Spans.Min(static span => span.PaintOrder));
        EnsureAppendOrderIsSafe(page, detected.Spans, replacementRequests);
        requests.AddRange(replacementRequests);
        byte[] output = ApplyStampRequests(removal.Bytes, requests, readOptions);
        IEnumerable<string> warnings = removal.Warnings.Concat(BuildSubstitutionWarnings(detected, style));
        if (fitWarning is not null) warnings = warnings.Append(fitWarning);
        return new TextMutationResult(output, detected.Spans.Count, warnings);
    }

    internal static TextMutationResult Move(byte[] pdf, PdfPageRegion source, double deltaX, double deltaY, PdfTextEditOptions? options, PdfLoadOptions? readOptions) {
        ValidateFinite(deltaX, nameof(deltaX));
        ValidateFinite(deltaY, nameof(deltaY));
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(source, "region");
        var sourceDocument = new TextEditSource(pdf, readOptions);
        source = TranslateRegionToSource(sourceDocument.VisualDocument, source);
        PdfTextEditOptions snapshot = (options ?? new PdfTextEditOptions()).Snapshot();
        PdfReadPage page = sourceDocument.VisualDocument.Pages[source.PageNumber - 1];
        EnsureRegionIsSafelyEditable(page, source, snapshot.AllowTextRenderingMode3);
        PdfRegionText detected = InspectSource(page, source, snapshot.AllowTextRenderingMode3);
        EnsureCompatibleRenderingModes(detected.Spans, snapshot.AllowTextRenderingMode3);
        if (detected.Spans.Count == 0 || detected.Text.Length == 0) return new TextMutationResult(pdf.ToArray(), 0, Array.Empty<string>());
        EnsureAppendOrderIsSafe(page, detected.Spans);
        TextRemovalResult removal = RemoveTextPreservingUnmatchedSpans(pdf, new[] { source.ToRedactionArea() }, readOptions,
            allowTextRenderingMode3: snapshot.AllowTextRenderingMode3, sourceDocument: sourceDocument.Document);
        var requests = new List<PdfStamper.TextStampRequest>(removal.Restamps);
        var movedRequests = new List<PdfStamper.TextStampRequest>();
        var warnings = new List<string>(removal.Warnings);
        PdfReadDocument styleDocument = sourceDocument.Document;
        for (int index = 0; index < detected.Spans.Count; index++) {
            PdfTextSpan span = detected.Spans[index];
            PdfRegionText spanRegion = BuildRegionText(new[] { span });
            PdfResolvedTextStyle style = ResolveStyle(snapshot, spanRegion, styleDocument.Pages[source.PageNumber - 1]);
            warnings.AddRange(BuildSubstitutionWarnings(spanRegion, style));
            AddStampLines(movedRequests, source.PageNumber, span.X + deltaX, span.Y + deltaY, span.RestampText, style, span.PaintOrder);
        }
        EnsureAppendOrderIsSafe(page, detected.Spans, movedRequests);
        requests.AddRange(movedRequests);
        byte[] output = ApplyStampRequests(removal.Bytes, requests, readOptions);
        return new TextMutationResult(output, detected.Spans.Count, warnings);
    }

    internal static TextMutationResult Replace(byte[] pdf, PdfTextMatch match, string text, PdfTextEditOptions? options, PdfLoadOptions? readOptions) {
        Guard.NotNull(match, nameof(match));
        Guard.NotNull(text, nameof(text));
        TextSearchHit hit = ResolveHit(pdf, match, readOptions);
        return ReplaceHits(pdf, new[] { hit }, text, options, readOptions);
    }

    internal static TextMutationResult Move(byte[] pdf, PdfTextMatch match, double deltaX, double deltaY, PdfTextEditOptions? options, PdfLoadOptions? readOptions) {
        Guard.NotNull(match, nameof(match));
        ValidateFinite(deltaX, nameof(deltaX));
        ValidateFinite(deltaY, nameof(deltaY));
        TextSearchHit hit = ResolveHit(pdf, match, readOptions);
        PdfTextSpan[] targetSpans = hit.Segments.Select(static segment => segment.Span).Distinct().ToArray();
        PdfTextEditOptions snapshot = (options ?? new PdfTextEditOptions()).Snapshot();
        EnsureCompatibleRenderingModes(targetSpans, snapshot.AllowTextRenderingMode3);
        EnsureAppendOrderIsSafe(pdf, hit.PageNumber, targetSpans, readOptions);

        PageSpanKey[] keys = targetSpans.Select(span => new PageSpanKey(hit.PageNumber, span)).ToArray();
        PdfRedactionArea[] areas = keys.Select(static key => {
            SpanBounds bounds = GetBounds(key.Span);
            return new PdfRedactionArea(key.PageNumber, bounds.X, bounds.Y, bounds.Width, bounds.Height);
        }).ToArray();
        TextRemovalResult removal = RemoveTextPreservingUnmatchedSpans(pdf, areas, readOptions, keys, allowTextRenderingMode3: snapshot.AllowTextRenderingMode3);
        var warnings = new List<string>(removal.Warnings);
        var requests = new List<PdfStamper.TextStampRequest>(removal.Restamps);
        var movedRequests = new List<PdfStamper.TextStampRequest>();
        PdfReadDocument styleDocument = PdfReadDocument.Open(pdf, readOptions);
        foreach (PdfTextSpan sourceSpan in targetSpans) {
            PdfRegionText detected = BuildRegionText(new[] { sourceSpan });
            PdfResolvedTextStyle sourceStyle = ResolveStyle(new PdfTextEditOptions(), detected, styleDocument.Pages[hit.PageNumber - 1]);
            PdfResolvedTextStyle movedStyle = ResolveStyle(snapshot, detected, styleDocument.Pages[hit.PageNumber - 1]);
            warnings.AddRange(BuildSubstitutionWarnings(detected, movedStyle));
            AddExactMoveRequests(
                movedRequests,
                hit.PageNumber,
                sourceSpan,
                hit.Segments.Where(segment => ReferenceEquals(segment.Span, sourceSpan)).ToArray(),
                deltaX,
                deltaY,
                sourceStyle,
                movedStyle);
        }
        EnsureAppendOrderIsSafe(pdf, hit.PageNumber, targetSpans, readOptions, movedRequests);
        requests.AddRange(movedRequests);
        return new TextMutationResult(ApplyStampRequests(removal.Bytes, requests, readOptions), 1, warnings);
    }

    internal static TextMutationResult ReplaceAll(byte[] pdf, string text, string replacement, PdfTextSearchOptions? searchOptions, PdfTextEditOptions? editOptions, PdfLoadOptions? readOptions) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(replacement, nameof(replacement));
        if (text.Length == 0) throw new ArgumentException("Search text cannot be empty.", nameof(text));
        IReadOnlyList<TextSearchHit> hits = FindHits(pdf, text, searchOptions, readOptions);
        if (hits.Count == 0) return new TextMutationResult(pdf.ToArray(), 0, Array.Empty<string>());
        return ReplaceHits(pdf, hits, replacement, editOptions, readOptions);
    }

    internal static TextMutationResult ReplaceSelected(byte[] pdf, string text, string replacement, int[] indexes,
        PdfTextSearchOptions? searchOptions, PdfTextEditOptions? editOptions, PdfLoadOptions? readOptions) {
        Guard.NotNull(replacement, nameof(replacement));
        IReadOnlyList<TextSearchHit> hits = FindHits(pdf, text, searchOptions, readOptions);
        if (indexes.Any(index => index < 0 || index >= hits.Count)) throw new ArgumentOutOfRangeException(nameof(indexes), "Selected occurrence indexes must identify current search results.");
        if (indexes.Distinct().Count() != indexes.Length) throw new ArgumentException("Duplicate occurrence indexes are not supported.", nameof(indexes));
        if (indexes.Length == 0) return new TextMutationResult(pdf.ToArray(), 0, Array.Empty<string>());
        return ReplaceHits(pdf, indexes.OrderBy(index => index).Select(index => hits[index]).ToArray(), replacement, editOptions, readOptions);
    }

    private static TextMutationResult ReplaceHits(byte[] pdf, IReadOnlyList<TextSearchHit> hits, string replacement, PdfTextEditOptions? editOptions, PdfLoadOptions? readOptions) {
        PdfTextEditOptions snapshot = (editOptions ?? new PdfTextEditOptions()).Snapshot();
        foreach (IGrouping<int, TextSearchHit> pageHits in hits.GroupBy(static hit => hit.PageNumber)) {
            foreach (TextSearchHit hit in pageHits) {
                EnsureCompatibleRenderingModes(
                    hit.Segments.Select(static segment => segment.Span).Distinct().ToArray(),
                    snapshot.AllowTextRenderingMode3);
            }
            PdfTextSpan[] targetSpans = pageHits.SelectMany(static hit => hit.Segments).Select(static segment => segment.Span).Distinct().ToArray();
            EnsureAppendOrderIsSafe(pdf, pageHits.Key, targetSpans, readOptions);
        }

        var rewrites = new Dictionary<PageSpanKey, List<SpanTextEdit>>();
        var fitWarnings = new List<string>();
        PdfReadDocument styleDocument = PdfReadDocument.Open(pdf, readOptions);
        for (int hitIndex = 0; hitIndex < hits.Count; hitIndex++) {
            TextSearchHit hit = hits[hitIndex];
            PdfResolvedTextStyle? fittedStyle = null;
            if (snapshot.RegionWidthPolicy != PdfTextRegionWidthPolicy.PreserveFontSize && replacement.Length > 0) {
                PdfResolvedTextStyle style = ResolveStyle(snapshot, BuildRegionText(new[] { hit.Segments[0].Span }), styleDocument.Pages[hit.PageNumber - 1]);
                fittedStyle = FitStyleToBaselineExtent(style, replacement, GetMatchedBaselineExtent(hit.Segments, hit.Lines[0]),
                    snapshot, out string? fitWarning);
                if (fitWarning is not null) fitWarnings.Add(fitWarning);
            }
            for (int segmentIndex = 0; segmentIndex < hit.Segments.Length; segmentIndex++) {
                TextSourceSegment segment = hit.Segments[segmentIndex];
                var key = new PageSpanKey(hit.PageNumber, segment.Span);
                if (!rewrites.TryGetValue(key, out List<SpanTextEdit>? edits)) {
                    edits = new List<SpanTextEdit>();
                    rewrites.Add(key, edits);
                }
                edits.Add(new SpanTextEdit(segment.Start, segment.Length, segmentIndex == 0 ? replacement : string.Empty, fittedStyle));
            }
        }
        IncludeTrailingFlowSpans(rewrites, hits, snapshot.AllowTextRenderingMode3);

        PdfRedactionArea[] areas = rewrites.Keys
            .Select(static key => {
                SpanBounds bounds = GetBounds(key.Span);
                return new PdfRedactionArea(key.PageNumber, bounds.X, bounds.Y, bounds.Width, bounds.Height);
            })
            .ToArray();
        TextRemovalResult removal = RemoveTextPreservingUnmatchedSpans(pdf, areas, readOptions, rewrites.Keys.ToArray(), allowTextRenderingMode3: snapshot.AllowTextRenderingMode3);
        var warnings = new List<string>(removal.Warnings.Concat(fitWarnings));
        var requests = new List<PdfStamper.TextStampRequest>(removal.Restamps);
        var rewrittenRequests = new List<PdfStamper.TextStampRequest>();
        var positioned = new List<PositionedRewrite>();
        foreach (KeyValuePair<PageSpanKey, List<SpanTextEdit>> rewrite in rewrites
            .OrderBy(static item => item.Key.PageNumber)
            .ThenByDescending(static item => item.Key.Span.Y)
            .ThenBy(static item => item.Key.Span.X)) {
            PdfTextSpan sourceSpan = rewrite.Key.Span;
            PdfRegionText detected = BuildRegionText(new[] { sourceSpan });
            PdfResolvedTextStyle sourceStyle = ResolveStyle(new PdfTextEditOptions(), detected, styleDocument.Pages[rewrite.Key.PageNumber - 1]);
            PdfResolvedTextStyle replacementStyle = ResolveStyle(snapshot, detected, styleDocument.Pages[rewrite.Key.PageNumber - 1]);
            PositionedTextFragment[] fragments = BuildPositionedFragments(sourceSpan, rewrite.Value, sourceStyle, replacementStyle);
            for (int fragmentIndex = 0; fragmentIndex < fragments.Length; fragmentIndex++) {
                warnings.AddRange(BuildSubstitutionWarnings(detected, fragments[fragmentIndex].Style));
            }
            positioned.Add(new PositionedRewrite(rewrite.Key.PageNumber, sourceSpan, fragments));
        }

        AddReflowedRewriteRequests(rewrittenRequests, positioned);
        foreach (IGrouping<int, PageSpanKey> pageRewrites in rewrites.Keys.GroupBy(static key => key.PageNumber)) {
            EnsureAppendOrderIsSafe(
                pdf,
                pageRewrites.Key,
                pageRewrites.Select(static key => key.Span).ToArray(),
                readOptions,
                rewrittenRequests.Where(request => request.PageNumber == pageRewrites.Key).ToArray());
        }
        requests.AddRange(rewrittenRequests);
        byte[] current = ApplyStampRequests(removal.Bytes, requests, readOptions);

        return new TextMutationResult(current, hits.Count, warnings);
    }

}

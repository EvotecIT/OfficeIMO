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
        region = TranslateRegionToSource(pdf, region, readOptions);
        PdfTextEditOptions snapshot = (options ?? new PdfTextEditOptions()).Snapshot();
        EnsureRegionIsSafelyEditable(pdf, region, snapshot.AllowTextRenderingMode3, readOptions);
        PdfRegionText detected = InspectSource(pdf, region, snapshot.AllowTextRenderingMode3, readOptions);
        EnsureCompatibleRenderingModes(detected.Spans, snapshot.AllowTextRenderingMode3);
        EnsureAppendOrderIsSafe(pdf, region.PageNumber, detected.Spans, readOptions);
        PdfResolvedTextStyle style = ResolveStyle(snapshot, detected);
        string replacementText = PreserveAuthoredEdgeWhitespace(detected.Spans, text);
        style = FitStyleToRegion(style, replacementText, region, snapshot, out string? fitWarning);
        TextRemovalResult removal = detected.Spans.Count == 0
            ? new TextRemovalResult(pdf.ToArray(), Array.Empty<string>(), Array.Empty<PdfStamper.TextStampRequest>())
            : RemoveTextPreservingUnmatchedSpans(pdf, new[] { region.ToRedactionArea() }, readOptions, allowTextRenderingMode3: snapshot.AllowTextRenderingMode3);
        var requests = new List<PdfStamper.TextStampRequest>(removal.Restamps);
        var replacementRequests = new List<PdfStamper.TextStampRequest>();
        if (text.Length > 0) AddStampLines(replacementRequests, region.PageNumber, detected.Spans.Count == 0 ? region.X : detected.BaselineX, detected.Spans.Count == 0 ? region.Top - style.FontSize : detected.BaselineY, replacementText, style, detected.Spans.Count == 0 ? double.MaxValue : detected.Spans.Min(static span => span.PaintOrder));
        EnsureAppendOrderIsSafe(pdf, region.PageNumber, detected.Spans, readOptions, replacementRequests);
        requests.AddRange(replacementRequests);
        byte[] output = ApplyStampRequests(removal.Bytes, requests, readOptions);
        IEnumerable<string> warnings = removal.Warnings.Concat(BuildSubstitutionWarnings(detected, style.Font));
        if (fitWarning is not null) warnings = warnings.Append(fitWarning);
        return new TextMutationResult(output, detected.Spans.Count, warnings);
    }

    internal static TextMutationResult Move(byte[] pdf, PdfPageRegion source, double deltaX, double deltaY, PdfTextEditOptions? options, PdfLoadOptions? readOptions) {
        ValidateFinite(deltaX, nameof(deltaX));
        ValidateFinite(deltaY, nameof(deltaY));
        source = TranslateRegionToSource(pdf, source, readOptions);
        PdfTextEditOptions snapshot = (options ?? new PdfTextEditOptions()).Snapshot();
        EnsureRegionIsSafelyEditable(pdf, source, snapshot.AllowTextRenderingMode3, readOptions);
        PdfRegionText detected = InspectSource(pdf, source, snapshot.AllowTextRenderingMode3, readOptions);
        EnsureCompatibleRenderingModes(detected.Spans, snapshot.AllowTextRenderingMode3);
        if (detected.Spans.Count == 0 || detected.Text.Length == 0) return new TextMutationResult(pdf.ToArray(), 0, Array.Empty<string>());
        EnsureAppendOrderIsSafe(pdf, source.PageNumber, detected.Spans, readOptions);
        TextRemovalResult removal = RemoveTextPreservingUnmatchedSpans(pdf, new[] { source.ToRedactionArea() }, readOptions, allowTextRenderingMode3: snapshot.AllowTextRenderingMode3);
        var requests = new List<PdfStamper.TextStampRequest>(removal.Restamps);
        var movedRequests = new List<PdfStamper.TextStampRequest>();
        var warnings = new List<string>(removal.Warnings);
        for (int index = 0; index < detected.Spans.Count; index++) {
            PdfTextSpan span = detected.Spans[index];
            PdfRegionText spanRegion = BuildRegionText(new[] { span });
            PdfResolvedTextStyle style = ResolveStyle(snapshot, spanRegion);
            warnings.AddRange(BuildSubstitutionWarnings(spanRegion, style.Font));
            AddStampLines(movedRequests, source.PageNumber, span.X + deltaX, span.Y + deltaY, span.RestampText, style, span.PaintOrder);
        }
        EnsureAppendOrderIsSafe(pdf, source.PageNumber, detected.Spans, readOptions, movedRequests);
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
        foreach (PdfTextSpan sourceSpan in targetSpans) {
            PdfRegionText detected = BuildRegionText(new[] { sourceSpan });
            PdfResolvedTextStyle sourceStyle = ResolveStyle(new PdfTextEditOptions(), detected);
            PdfResolvedTextStyle movedStyle = ResolveStyle(snapshot, detected);
            warnings.AddRange(BuildSubstitutionWarnings(detected, movedStyle.Font));
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
        for (int hitIndex = 0; hitIndex < hits.Count; hitIndex++) {
            TextSearchHit hit = hits[hitIndex];
            PdfResolvedTextStyle? fittedStyle = null;
            if (snapshot.RegionWidthPolicy != PdfTextRegionWidthPolicy.PreserveFontSize && replacement.Length > 0) {
                PdfResolvedTextStyle style = ResolveStyle(snapshot, BuildRegionText(new[] { hit.Segments[0].Span }));
                fittedStyle = FitStyleToRegion(style, replacement,
                    new PdfPageRegion(hit.PageNumber, hit.Match.X, hit.Match.Y, hit.Match.Width, hit.Match.Height),
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
            PdfResolvedTextStyle sourceStyle = ResolveStyle(new PdfTextEditOptions(), detected);
            PdfResolvedTextStyle replacementStyle = ResolveStyle(snapshot, detected);
            PositionedTextFragment[] fragments = BuildPositionedFragments(sourceSpan, rewrite.Value, sourceStyle, replacementStyle);
            for (int fragmentIndex = 0; fragmentIndex < fragments.Length; fragmentIndex++) {
                warnings.AddRange(BuildSubstitutionWarnings(detected, fragments[fragmentIndex].Style.Font));
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

using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Canonical extraction, removal, and stamping coordinator for existing-page text edits.</summary>
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
        double baselineY = region.Top - style.FontSize;
        byte[] output = StampLines(pdf, region.PageNumber, region.X, baselineY, text, style, readOptions);
        return new TextMutationResult(output, 1, Array.Empty<string>());
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
        TextRemovalResult removal = detected.Spans.Count == 0
            ? new TextRemovalResult(pdf.ToArray(), Array.Empty<string>(), Array.Empty<PdfStamper.TextStampRequest>())
            : RemoveTextPreservingUnmatchedSpans(pdf, new[] { region.ToRedactionArea() }, readOptions, allowTextRenderingMode3: snapshot.AllowTextRenderingMode3);
        var requests = new List<PdfStamper.TextStampRequest>(removal.Restamps);
        var replacementRequests = new List<PdfStamper.TextStampRequest>();
        if (text.Length > 0) AddStampLines(replacementRequests, region.PageNumber, detected.Spans.Count == 0 ? region.X : detected.BaselineX, detected.Spans.Count == 0 ? region.Top - style.FontSize : detected.BaselineY, PreserveAuthoredEdgeWhitespace(detected.Spans, text), style, detected.Spans.Count == 0 ? double.MaxValue : detected.Spans.Min(static span => span.PaintOrder));
        EnsureAppendOrderIsSafe(pdf, region.PageNumber, detected.Spans, readOptions, replacementRequests);
        requests.AddRange(replacementRequests);
        byte[] output = ApplyStampRequests(removal.Bytes, requests, readOptions);
        return new TextMutationResult(output, detected.Spans.Count, removal.Warnings.Concat(BuildSubstitutionWarnings(detected, style.Font)));
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
        for (int hitIndex = 0; hitIndex < hits.Count; hitIndex++) {
            TextSearchHit hit = hits[hitIndex];
            for (int segmentIndex = 0; segmentIndex < hit.Segments.Length; segmentIndex++) {
                TextSourceSegment segment = hit.Segments[segmentIndex];
                var key = new PageSpanKey(hit.PageNumber, segment.Span);
                if (!rewrites.TryGetValue(key, out List<SpanTextEdit>? edits)) {
                    edits = new List<SpanTextEdit>();
                    rewrites.Add(key, edits);
                }
                edits.Add(new SpanTextEdit(segment.Start, segment.Length, segmentIndex == 0 ? replacement : string.Empty));
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
        var warnings = new List<string>(removal.Warnings);
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

    internal static byte[] RemoveExactContentSafetySpans(
        byte[] pdf,
        IReadOnlyList<(int PageNumber, PdfTextSpan Span)> targets,
        PdfLoadOptions? readOptions) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(targets);
#else
        if (targets == null) throw new ArgumentNullException(nameof(targets));
#endif
        if (targets.Count == 0) return pdf.ToArray();
        PageSpanKey[] keys = targets.Select(item => new PageSpanKey(item.PageNumber, item.Span)).ToArray();
        PdfRedactionArea[] areas = keys.Select(key => {
            SpanBounds bounds = GetBounds(key.Span);
            return new PdfRedactionArea(key.PageNumber, bounds.X, bounds.Y, bounds.Width, bounds.Height, "content-safety");
        }).ToArray();
        TextRemovalResult removal = RemoveTextPreservingUnmatchedSpans(
            pdf,
            areas,
            readOptions,
            keys,
            allowInvisibleTargetRemoval: true);
        return ApplyStampRequests(removal.Bytes, new List<PdfStamper.TextStampRequest>(removal.Restamps), readOptions);
    }

    internal static byte[] MutateExactContentSafetySpans(
        byte[] pdf,
        IReadOnlyList<(int PageNumber, PdfTextSpan Span)> removals,
        IReadOnlyList<(int PageNumber, PdfTextSpan Span, int Start, int Length)> textEdits,
        PdfLoadOptions? readOptions) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(removals);
        ArgumentNullException.ThrowIfNull(textEdits);
#else
        if (removals == null) throw new ArgumentNullException(nameof(removals));
        if (textEdits == null) throw new ArgumentNullException(nameof(textEdits));
#endif
        if (textEdits.Count == 0) return RemoveExactContentSafetySpans(pdf, removals, readOptions);
        var rewrites = new Dictionary<PageSpanKey, List<SpanTextEdit>>();
        foreach (var edit in textEdits) {
            if (edit.Start < 0 || edit.Length <= 0 || edit.Start > edit.Span.Text.Length - edit.Length) {
                throw new InvalidOperationException("A selected PDF Unicode range no longer matches the inspected text span.");
            }
            var key = new PageSpanKey(edit.PageNumber, edit.Span);
            if (!rewrites.TryGetValue(key, out List<SpanTextEdit>? ranges)) {
                ranges = new List<SpanTextEdit>();
                rewrites.Add(key, ranges);
            }
            ranges.Add(new SpanTextEdit(edit.Start, edit.Length, string.Empty));
        }
        PageSpanKey[] removalKeys = removals.Select(item => new PageSpanKey(item.PageNumber, item.Span)).ToArray();
        PageSpanKey[] allKeys = removalKeys.Concat(rewrites.Keys).Distinct().ToArray();
        PdfRedactionArea[] areas = allKeys.Select(key => {
            SpanBounds bounds = GetBounds(key.Span);
            return new PdfRedactionArea(key.PageNumber, bounds.X, bounds.Y, bounds.Width, bounds.Height, "content-safety");
        }).ToArray();
        TextRemovalResult removal = RemoveTextPreservingUnmatchedSpans(pdf, areas, readOptions, allKeys, allowInvisibleTargetRemoval: true);
        var requests = new List<PdfStamper.TextStampRequest>(removal.Restamps);
        var rewrittenRequests = new List<PdfStamper.TextStampRequest>();
        var positioned = new List<PositionedRewrite>();
        foreach (KeyValuePair<PageSpanKey, List<SpanTextEdit>> rewrite in rewrites
            .OrderBy(item => item.Key.PageNumber)
            .ThenByDescending(item => item.Key.Span.Y)
            .ThenBy(item => item.Key.Span.X)) {
            PdfTextSpan sourceSpan = rewrite.Key.Span;
            PdfRegionText detected = BuildRegionText(new[] { sourceSpan });
            PdfResolvedTextStyle style = ResolveStyle(new PdfTextEditOptions(), detected);
            PositionedTextFragment[] fragments = BuildPositionedFragments(sourceSpan, rewrite.Value, style, style);
            positioned.Add(new PositionedRewrite(rewrite.Key.PageNumber, sourceSpan, fragments));
        }
        AddReflowedRewriteRequests(rewrittenRequests, positioned);
        foreach (IGrouping<int, PageSpanKey> pageRewrites in rewrites.Keys.GroupBy(key => key.PageNumber)) {
            EnsureAppendOrderIsSafe(
                pdf,
                pageRewrites.Key,
                pageRewrites.Select(key => key.Span).ToArray(),
                readOptions,
                rewrittenRequests.Where(request => request.PageNumber == pageRewrites.Key).ToArray());
        }
        requests.AddRange(rewrittenRequests);
        return ApplyStampRequests(removal.Bytes, requests, readOptions);
    }

    private static TextRemovalResult RemoveTextPreservingUnmatchedSpans(
        byte[] pdf,
        IReadOnlyList<PdfRedactionArea> areas,
        PdfLoadOptions? readOptions,
        IReadOnlyList<PageSpanKey>? exactTargets = null,
        bool allowInvisibleTargetRemoval = false,
        bool allowTextRenderingMode3 = false) {
        PdfReadDocument before = PdfReadDocument.Open(pdf, readOptions);
        int[] affectedPages = areas.Select(static area => area.PageNumber).Distinct().ToArray();
        var original = new List<PageTextSpanSnapshot>();
        for (int index = 0; index < affectedPages.Length; index++) {
            int pageNumber = affectedPages[index];
            ValidatePage(pageNumber, before.Pages.Count, nameof(areas));
            IReadOnlyList<PdfTextSpan> spans = before.Pages[pageNumber - 1].GetTextSpans(includeArtifactText: true);
            for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++) {
                PdfTextSpan span = spans[spanIndex];
                SpanBounds bounds = GetBounds(span);
                bool targeted = exactTargets is null
                    ? areas.Any(area => area.PageNumber == pageNumber && Intersects(area, bounds))
                    : exactTargets.Any(target => target.PageNumber == pageNumber && SameTargetSourceSpan(span, target.Span));
                if (targeted && !allowInvisibleTargetRemoval && !IsSafelyEditableSpan(span, allowTextRenderingMode3)) {
                    throw new NotSupportedException("The selected region contains invisible or clipped text whose rendering state cannot be recreated safely.");
                }
                original.Add(new PageTextSpanSnapshot(pageNumber, span, targeted));
            }
        }

        IReadOnlyList<PdfRedactionArea> removalAreas = exactTargets is null
            ? areas
            : exactTargets.Select(static target => {
                SpanBounds bounds = GetBounds(target.Span);
                return new PdfRedactionArea(target.PageNumber, bounds.X, bounds.Y, bounds.Width, bounds.Height)
                    .WithTextRenderingMode(target.Span.TextRenderingMode);
            }).ToArray();
        byte[] removed = PdfRedactionApplier.RemoveTextInAreas(pdf, removalAreas, readOptions: readOptions);
        PdfLoadOptions afterReadOptions = PdfLoadOptions.WithMinimumInputBytes(readOptions, removed.LongLength);
        PdfReadDocument after = PdfReadDocument.Open(removed, afterReadOptions);
        var remainingByPage = affectedPages.ToDictionary(
            static page => page,
            page => after.Pages[page - 1].GetTextSpans(includeArtifactText: true).ToList());
        var missing = new List<PageTextSpanSnapshot>();
        for (int index = 0; index < original.Count; index++) {
            PageTextSpanSnapshot candidate = original[index];
            if (candidate.Targeted) continue;
            List<PdfTextSpan> remaining = remainingByPage[candidate.PageNumber];
            int matchIndex = remaining.FindIndex(span => SameSurvivingSourceSpan(span, candidate.Span));
            if (matchIndex >= 0) remaining.RemoveAt(matchIndex);
            else missing.Add(candidate);
        }

        var warnings = new List<string>();
        var restamps = new List<PdfStamper.TextStampRequest>();
        foreach (IGrouping<int, PageTextSpanSnapshot> pageMissing in missing.GroupBy(static snapshot => snapshot.PageNumber)) {
            if (before.Pages[pageMissing.Key - 1].WouldAppendingTextChangeVisibleStacking(pageMissing.Select(static snapshot => snapshot.Span).ToArray())) {
                throw new NotSupportedException("The text edit would change the visible stacking order while restoring adjacent source text.");
            }
        }
        for (int index = 0; index < missing.Count; index++) {
            PageTextSpanSnapshot snapshot = missing[index];
            if (!IsSafelyEditableSpan(snapshot.Span, allowTextRenderingMode3)) {
                throw new NotSupportedException("The text edit would require recreating invisible or clipped source text without its original rendering state.");
            }
            PdfRegionText detected = BuildRegionText(new[] { snapshot.Span });
            PdfResolvedTextStyle style = ResolveStyle(new PdfTextEditOptions(), detected);
            warnings.AddRange(BuildSubstitutionWarnings(detected, style.Font));
            AddStampLines(restamps, snapshot.PageNumber, snapshot.Span.X, snapshot.Span.Y, snapshot.Span.RestampText, style, snapshot.Span.PaintOrder);
        }
        return new TextRemovalResult(removed, warnings, restamps);
    }

    private static PdfRegionText BuildRegionText(PdfTextSpan[] spans) {
        if (spans.Length == 0) return new PdfRegionText(string.Empty, Array.Empty<PdfTextSpan>(), PdfStandardFont.Helvetica, null, 12D, PdfColor.Black, 0D, 0D, 0D);
        PdfTextSpan dominant = spans
            .GroupBy(static span => new SpanStyleKey(span.BaseFont ?? span.FontResource, Math.Round(EffectiveFontSize(span), 2), Math.Round(span.RotationDegrees, 2), span.Color))
            .OrderByDescending(static group => group.Sum(static span => Math.Max(1, span.Text.Length)))
            .First()
            .First();
        List<TextLayoutEngine.TextLine> layoutLines = BuildSearchLines(spans);
        PdfTextSpan[] ordered = layoutLines.SelectMany(static line => line.Spans).ToArray();
        var builder = new System.Text.StringBuilder();
        for (int lineIndex = 0; lineIndex < layoutLines.Count; lineIndex++) {
            if (lineIndex > 0) builder.Append('\n');
            builder.Append(new TextSearchUnit(layoutLines[lineIndex].Spans).Text);
        }
        return new PdfRegionText(builder.ToString(), ordered, ResolveStandardFont(dominant.BaseFont), dominant.BaseFont, EffectiveFontSize(dominant), ToPdfColor(dominant.Color), dominant.RotationDegrees, ordered[0].X, ordered[0].Y);
    }

    private static PdfRegionText OffsetRegionText(PdfRegionText value, double deltaX, double deltaY) {
        if (value.Spans.Count == 0) return value;
        return new PdfRegionText(
            value.Text,
            value.Spans.Select(span => span.WithOffset(deltaX, deltaY)).ToArray(),
            value.SuggestedFont,
            value.SourceFont,
            value.FontSize,
            value.Color,
            value.RotationDegrees,
            value.BaselineX + deltaX,
            value.BaselineY + deltaY);
    }

    private static byte[] StampLines(byte[] pdf, int pageNumber, double x, double baselineY, string text, PdfResolvedTextStyle style, PdfLoadOptions? readOptions) {
        var requests = new List<PdfStamper.TextStampRequest>();
        AddStampLines(requests, pageNumber, x, baselineY, text, style);
        return ApplyStampRequests(pdf, requests, readOptions);
    }

    private static void AddStampLines(List<PdfStamper.TextStampRequest> requests, int pageNumber, double x, double baselineY, string text, PdfResolvedTextStyle style, double paintOrder = double.MaxValue) {
        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        double radians = style.RotationDegrees * Math.PI / 180D;
        double normalX = Math.Sin(radians) * style.FontSize * 1.15D;
        double normalY = -Math.Cos(radians) * style.FontSize * 1.15D;
        for (int index = 0; index < lines.Length; index++) {
            if (lines[index].Length == 0) continue;
            requests.Add(new PdfStamper.TextStampRequest(
                pageNumber,
                lines[index],
                x + normalX * index,
                baselineY + normalY * index,
                style.Font,
                style.FontSize,
                style.Color,
                style.RotationDegrees,
                paintOrder + (index * 0.0000001D),
                style.TextRenderingMode));
        }
    }

    private static bool HasRenderableTextLine(string text) {
        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int index = 0; index < lines.Length; index++) if (lines[index].Length > 0) return true;
        return false;
    }

    private static byte[] ApplyStampRequests(byte[] pdf, List<PdfStamper.TextStampRequest> requests, PdfLoadOptions? readOptions) =>
        requests.Count == 0
            ? pdf
            : PdfStamper.StampTextBatch(pdf, requests, PdfLoadOptions.WithMinimumInputBytes(readOptions, pdf.LongLength));

    private static void AddReflowedRewriteRequests(List<PdfStamper.TextStampRequest> requests, IReadOnlyList<PositionedRewrite> rewrites) {
        foreach (IGrouping<RewriteLineKey, PositionedRewrite> rotationGroup in rewrites.GroupBy(static rewrite => RewriteLineKey.Create(rewrite))) {
            PositionedRewrite[] candidates = rotationGroup.ToArray();
            bool rightToLeft = UsesRightToLeftReadingOrder(candidates.Select(static rewrite => rewrite.Source).ToArray());
            PositionedRewrite[] ordered = rightToLeft
                ? candidates.OrderByDescending(static rewrite => NormalPosition(rewrite.Source))
                    .ThenByDescending(static rewrite => BaselinePosition(rewrite.Source)).ToArray()
                : candidates.OrderByDescending(static rewrite => NormalPosition(rewrite.Source))
                    .ThenBy(static rewrite => BaselinePosition(rewrite.Source)).ToArray();
            var line = new List<PositionedRewrite>();
            double normal = 0D;
            for (int index = 0; index < ordered.Length; index++) {
                PositionedRewrite rewrite = ordered[index];
                double rewriteNormal = NormalPosition(rewrite.Source);
                double tolerance = Math.Max(0.5D, Math.Min(2.5D, EffectiveFontSize(rewrite.Source) * 0.6D));
                bool startsIndependentFlow = false;
                if (line.Count > 0 && Math.Abs(rewriteNormal - normal) <= tolerance) {
                    PositionedRewrite previous = line[line.Count - 1];
                    double gap = rightToLeft
                        ? BaselinePosition(previous.Source) - Math.Abs(previous.Source.Advance) - BaselinePosition(rewrite.Source)
                        : BaselinePosition(rewrite.Source) -
                            (BaselinePosition(previous.Source) + Math.Abs(previous.Source.Advance));
                    startsIndependentFlow = IsIndependentFlowGap(gap, previous.Source, rewrite.Source);
                }
                if (line.Count > 0 && (Math.Abs(rewriteNormal - normal) > tolerance || startsIndependentFlow)) {
                    AddReflowedLineRequests(requests, line);
                    line.Clear();
                }
                normal = line.Count == 0 ? rewriteNormal : ((normal * line.Count) + rewriteNormal) / (line.Count + 1);
                line.Add(rewrite);
            }
            if (line.Count > 0) AddReflowedLineRequests(requests, line);
        }
    }

    private static void IncludeTrailingFlowSpans(
        Dictionary<PageSpanKey, List<SpanTextEdit>> rewrites,
        IReadOnlyList<TextSearchHit> hits,
        bool allowTextRenderingMode3) {
        foreach (IGrouping<int, TextSearchHit> pageHits in hits.GroupBy(static hit => hit.PageNumber)) {
            foreach (IGrouping<PdfTextSpan[], TextSearchHit> lineHits in pageHits.GroupBy(
                         static hit => hit.LineSpans,
                         TextSpanArrayReferenceComparer.Instance)) {
                PdfTextSpan[] lineSpans = lineHits.Key;
                var ordinals = new Dictionary<PdfTextSpan, int>();
                for (int lineIndex = 0; lineIndex < lineSpans.Length; lineIndex++) {
                    ordinals[lineSpans[lineIndex]] = lineIndex;
                }

                int firstTrailingIndex = lineSpans.Length;
                foreach (TextSearchHit hit in lineHits) {
                    int lastMatchedIndex = -1;
                    for (int segmentIndex = 0; segmentIndex < hit.Segments.Length; segmentIndex++) {
                        if (ordinals.TryGetValue(hit.Segments[segmentIndex].Span, out int ordinal)) {
                            lastMatchedIndex = Math.Max(lastMatchedIndex, ordinal);
                        }
                    }
                    if (lastMatchedIndex >= 0) firstTrailingIndex = Math.Min(firstTrailingIndex, lastMatchedIndex + 1);
                }

                for (int lineIndex = firstTrailingIndex; lineIndex < lineSpans.Length; lineIndex++) {
                    var key = new PageSpanKey(pageHits.Key, lineSpans[lineIndex]);
                    if (!rewrites.ContainsKey(key)) {
                        rewrites.Add(key, new List<SpanTextEdit>());
                    }
                }

                EnsureCompatibleRenderingModes(
                    lineSpans.Where(span => rewrites.ContainsKey(new PageSpanKey(pageHits.Key, span))).ToArray(),
                    allowTextRenderingMode3);
            }
        }
    }

    private static void AddReflowedLineRequests(List<PdfStamper.TextStampRequest> requests, List<PositionedRewrite> line) {
        bool rightToLeft = UsesRightToLeftReadingOrder(line.Select(static rewrite => rewrite.Source).ToArray());
        PositionedRewrite[] ordered = rightToLeft
            ? line.OrderByDescending(static rewrite => BaselinePosition(rewrite.Source)).ToArray()
            : line.OrderBy(static rewrite => BaselinePosition(rewrite.Source)).ToArray();
        double offsetX = 0D;
        double offsetY = 0D;
        for (int index = 0; index < ordered.Length; index++) {
            PositionedRewrite rewrite = ordered[index];
            double radians = rewrite.Source.RotationDegrees * Math.PI / 180D;
            double ux = Math.Cos(radians);
            double uy = Math.Sin(radians);
            double x = rewrite.Source.X + offsetX;
            double y = rewrite.Source.Y + offsetY;
            double cursorX = x;
            double cursorY = y;
            for (int fragmentIndex = 0; fragmentIndex < rewrite.Fragments.Length; fragmentIndex++) {
                PositionedTextFragment fragment = rewrite.Fragments[fragmentIndex];
                if (fragment.Text.Length == 0) continue;
                AddStampLines(requests, rewrite.PageNumber, cursorX, cursorY, fragment.Text, fragment.Style, rewrite.Source.PaintOrder + (fragmentIndex * 0.00000001D));
                AdvanceFragmentCursor(fragment, ref cursorX, ref cursorY);
            }
            double sourceEndX = rewrite.Source.X + ux * Math.Abs(rewrite.Source.Advance);
            double sourceEndY = rewrite.Source.Y + uy * Math.Abs(rewrite.Source.Advance);
            double deltaX = cursorX - sourceEndX;
            double deltaY = cursorY - sourceEndY;
            double baselineDelta = deltaX * ux + deltaY * uy;
            double normalDeltaX = deltaX - baselineDelta * ux;
            double normalDeltaY = deltaY - baselineDelta * uy;
            double baselineDirection = rightToLeft ? -1D : 1D;
            offsetX = normalDeltaX + baselineDirection * baselineDelta * ux;
            offsetY = normalDeltaY + baselineDirection * baselineDelta * uy;
        }
    }

    private static void AdvanceFragmentCursor(PositionedTextFragment fragment, ref double cursorX, ref double cursorY) {
        string[] lines = fragment.Text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        int finalLineIndex = lines.Length - 1;

        double radians = fragment.Style.RotationDegrees * Math.PI / 180D;
        double ux = Math.Cos(radians);
        double uy = Math.Sin(radians);
        double normalX = Math.Sin(radians) * fragment.Style.FontSize * 1.15D;
        double normalY = -Math.Cos(radians) * fragment.Style.FontSize * 1.15D;
        double finalWidth = PdfWriter.EstimateSimpleTextWidth(lines[finalLineIndex], fragment.Style.Font, fragment.Style.FontSize);
        cursorX += normalX * finalLineIndex + ux * finalWidth;
        cursorY += normalY * finalLineIndex + uy * finalWidth;
    }

    private static PdfResolvedTextStyle ResolveStyle(PdfTextEditOptions options, PdfRegionText? detected) => new PdfResolvedTextStyle(
        options.Font ?? detected?.SuggestedFont ?? PdfStandardFont.Helvetica,
        options.FontSize ?? detected?.FontSize ?? 12D,
        options.Color ?? detected?.Color ?? PdfColor.Black,
        options.RotationDegrees ?? detected?.RotationDegrees ?? 0D,
        detected?.Spans.Count > 0 && detected.Spans.All(static span => span.TextRenderingMode == 3) ? 3 : 0);

    private static string[] BuildSubstitutionWarnings(PdfRegionText detected, PdfStandardFont targetFont) {
        string source = StripSubsetPrefix(detected.SourceFont);
        if (source.Length == 0 || string.Equals(source, targetFont.ToBaseFontName(), StringComparison.OrdinalIgnoreCase)) return Array.Empty<string>();
        return new[] { "The source font '" + source + "' is not reused by the dependency-free text editor; replacement text uses '" + targetFont.ToBaseFontName() + "'. Metrics and letterforms can differ." };
    }

    private static PdfStandardFont ResolveStandardFont(string? baseFont) {
        string name = StripSubsetPrefix(baseFont).ToLowerInvariant();
        bool bold = name.Contains("bold") || name.Contains("black") || name.Contains("heavy") || name.Contains("semibold");
        bool italic = name.Contains("italic") || name.Contains("oblique");
        bool sansSerif = name.Contains("sansserif") || name.Contains("sans-serif") || name.Contains("sans serif");
        if (name.Contains("times") || (!sansSerif && name.Contains("serif")) || name.Contains("roman") || name.Contains("georgia")) {
            return bold && italic ? PdfStandardFont.TimesBoldItalic : bold ? PdfStandardFont.TimesBold : italic ? PdfStandardFont.TimesItalic : PdfStandardFont.TimesRoman;
        }
        if (name.Contains("courier") || name.Contains("mono") || name.Contains("consol")) {
            return bold && italic ? PdfStandardFont.CourierBoldOblique : bold ? PdfStandardFont.CourierBold : italic ? PdfStandardFont.CourierOblique : PdfStandardFont.Courier;
        }
        return bold && italic ? PdfStandardFont.HelveticaBoldOblique : bold ? PdfStandardFont.HelveticaBold : italic ? PdfStandardFont.HelveticaOblique : PdfStandardFont.Helvetica;
    }

    private static string StripSubsetPrefix(string? value) {
        string name = (value ?? string.Empty).Trim().TrimStart('/');
        return name.Length > 7 && name[6] == '+' ? name.Substring(7) : name;
    }

    private static PdfColor ToPdfColor(OfficeColor? color) => color.HasValue ? PdfColor.FromOfficeColor(color.Value) : PdfColor.Black;
    private static double EffectiveFontSize(PdfTextSpan span) => span.RestampFontSize > 0D && !double.IsNaN(span.RestampFontSize) && !double.IsInfinity(span.RestampFontSize) ? span.RestampFontSize : 12D;

    private static SpanBounds GetBounds(PdfTextSpan span) => SliceBounds(span, 0, Math.Max(1, span.Text.Length));

    private static SpanBounds GetCombinedSegmentBounds(IReadOnlyList<TextSourceSegment> segments) {
        SpanBounds combined = SliceBounds(segments[0].Span, segments[0].Start, segments[0].Length);
        double left = combined.X;
        double bottom = combined.Y;
        double right = combined.X + combined.Width;
        double top = combined.Y + combined.Height;
        for (int index = 1; index < segments.Count; index++) {
            TextSourceSegment segment = segments[index];
            SpanBounds bounds = SliceBounds(segment.Span, segment.Start, segment.Length);
            left = Math.Min(left, bounds.X);
            bottom = Math.Min(bottom, bounds.Y);
            right = Math.Max(right, bounds.X + bounds.Width);
            top = Math.Max(top, bounds.Y + bounds.Height);
        }
        return new SpanBounds(left, bottom, Math.Max(0.1D, right - left), Math.Max(0.1D, top - bottom));
    }

    private static SpanBounds SliceBounds(PdfTextSpan span, int start, int length) {
        double advance = Math.Max(Math.Abs(span.Advance), EffectiveFontSize(span) * Math.Max(1, span.Text.Length) * 0.45D);
        double offset;
        double sliceAdvance;
        if (PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries)) {
            int startIndex = Math.Min(start, boundaries.Length - 1);
            int endIndex = Math.Min(start + length, boundaries.Length - 1);
            offset = boundaries[startIndex];
            sliceAdvance = boundaries[endIndex] - offset;
        } else {
            double textLength = Math.Max(1, span.Text.Length);
            offset = advance * start / textLength;
            sliceAdvance = advance * length / textLength;
        }
        double radians = span.RotationDegrees * Math.PI / 180D;
        double ux = Math.Cos(radians);
        double uy = Math.Sin(radians);
        double nx = -uy;
        double ny = ux;
        double fontSize = EffectiveFontSize(span);
        double x0 = span.X + ux * offset - nx * fontSize * 0.25D;
        double y0 = span.Y + uy * offset - ny * fontSize * 0.25D;
        double x1 = span.X + ux * (offset + sliceAdvance) - nx * fontSize * 0.25D;
        double y1 = span.Y + uy * (offset + sliceAdvance) - ny * fontSize * 0.25D;
        double x2 = span.X + ux * offset + nx * fontSize * 0.8D;
        double y2 = span.Y + uy * offset + ny * fontSize * 0.8D;
        double x3 = span.X + ux * (offset + sliceAdvance) + nx * fontSize * 0.8D;
        double y3 = span.Y + uy * (offset + sliceAdvance) + ny * fontSize * 0.8D;
        double left = Math.Min(Math.Min(x0, x1), Math.Min(x2, x3));
        double right = Math.Max(Math.Max(x0, x1), Math.Max(x2, x3));
        double bottom = Math.Min(Math.Min(y0, y1), Math.Min(y2, y3));
        double top = Math.Max(Math.Max(y0, y1), Math.Max(y2, y3));
        return new SpanBounds(left, bottom, Math.Max(0.1D, right - left), Math.Max(0.1D, top - bottom));
    }

    private static bool Intersects(PdfPageRegion region, SpanBounds bounds) =>
        bounds.X < region.Right &&
        bounds.X + bounds.Width > region.X &&
        bounds.Y < region.Top &&
        bounds.Y + bounds.Height > region.Y;

    private static bool Intersects(PdfRedactionArea region, SpanBounds bounds) =>
        bounds.X < region.X + region.Width &&
        bounds.X + bounds.Width > region.X &&
        bounds.Y < region.Y + region.Height &&
        bounds.Y + bounds.Height > region.Y;

    private static bool SameTextPlacement(PdfTextSpan left, PdfTextSpan right) =>
        string.Equals(left.Text, right.Text, StringComparison.Ordinal) &&
        Math.Abs(left.X - right.X) <= 0.2D &&
        Math.Abs(left.Y - right.Y) <= 0.2D &&
        Math.Abs(EffectiveFontSize(left) - EffectiveFontSize(right)) <= 0.2D &&
        Math.Abs(left.RotationDegrees - right.RotationDegrees) <= 0.2D;

    private static bool SameTargetSourceSpan(PdfTextSpan left, PdfTextSpan right) =>
        SameSurvivingSourceSpan(left, right) &&
        Math.Abs(left.PaintOrder - right.PaintOrder) <= 0.0001D;

    private static bool SameSurvivingSourceSpan(PdfTextSpan left, PdfTextSpan right) =>
        SameTextPlacement(left, right) &&
        Math.Abs(left.Advance - right.Advance) <= 0.2D &&
        string.Equals(left.FontResource, right.FontResource, StringComparison.Ordinal) &&
        string.Equals(left.BaseFont, right.BaseFont, StringComparison.Ordinal) &&
        Nullable.Equals(left.Color, right.Color) &&
        left.IsVisible == right.IsVisible &&
        left.TextRenderingMode == right.TextRenderingMode &&
        left.CanRestamp == right.CanRestamp &&
        string.Equals(left.RestampText, right.RestampText, StringComparison.Ordinal) &&
        Math.Abs(left.RestampFontSize - right.RestampFontSize) <= 0.2D &&
        SameAdvances(left.CharacterAdvances, right.CharacterAdvances);

    private static bool SameAdvances(IReadOnlyList<double>? left, IReadOnlyList<double>? right) {
        if (ReferenceEquals(left, right)) return true;
        if (left == null || right == null || left.Count != right.Count) return false;
        for (int index = 0; index < left.Count; index++) {
            if (Math.Abs(left[index] - right[index]) > 0.01D) return false;
        }
        return true;
    }

    private static PositionedTextFragment[] BuildPositionedFragments(
        PdfTextSpan sourceSpan,
        IReadOnlyList<SpanTextEdit> edits,
        PdfResolvedTextStyle sourceStyle,
        PdfResolvedTextStyle replacementStyle) {
        string source = sourceSpan.Text;
        string authored = TrimAuthoredEdgeWhitespace(sourceSpan.RestampText);
        int[]? authoredBoundaries = TryBuildAuthoredBoundaryMap(source, authored);
        if (authoredBoundaries == null) {
            authored = source;
            authoredBoundaries = BuildIdentityBoundaryMap(source.Length);
        }
        SpanTextEdit[] ordered = edits.OrderBy(static edit => edit.Start).ToArray();
        var fragments = new List<PositionedTextFragment>();
        AddPositionedFragment(fragments, GetLeadingWhitespace(sourceSpan.RestampText), sourceStyle);
        int cursor = 0;
        for (int index = 0; index < ordered.Length; index++) {
            SpanTextEdit edit = ordered[index];
            if (edit.Start < cursor) continue;
            if (edit.Start > cursor) {
                int authoredStart = authoredBoundaries[cursor];
                int authoredEnd = authoredBoundaries[edit.Start];
                AddPositionedFragment(fragments, authored.Substring(authoredStart, authoredEnd - authoredStart), sourceStyle);
            }
            if (edit.Replacement.Length > 0) AddPositionedFragment(fragments, edit.Replacement, replacementStyle);
            cursor = edit.Start + edit.Length;
        }
        if (cursor < source.Length) {
            int authoredStart = authoredBoundaries[cursor];
            AddPositionedFragment(fragments, authored.Substring(authoredStart), sourceStyle);
        }
        AddPositionedFragment(fragments, GetTrailingWhitespace(sourceSpan.RestampText), sourceStyle);
        return fragments.ToArray();
    }

    private static string PreserveAuthoredEdgeWhitespace(IReadOnlyList<PdfTextSpan> spans, string replacement) {
        if (spans.Count == 0) return replacement;
        return GetLeadingWhitespace(spans[0].RestampText) + replacement + GetTrailingWhitespace(spans[spans.Count - 1].RestampText);
    }

    private static string GetLeadingWhitespace(string value) {
        int length = 0;
        while (length < value.Length && char.IsWhiteSpace(value[length])) length++;
        return length == 0 ? string.Empty : value.Substring(0, length);
    }

    private static string GetTrailingWhitespace(string value) {
        int start = value.Length;
        while (start > 0 && char.IsWhiteSpace(value[start - 1])) start--;
        return start == value.Length ? string.Empty : value.Substring(start);
    }

    private static void AddPositionedFragment(List<PositionedTextFragment> fragments, string text, PdfResolvedTextStyle style) {
        if (text.Length == 0) return;
        if (fragments.Count > 0 && fragments[fragments.Count - 1].Style.Equals(style)) {
            PositionedTextFragment previous = fragments[fragments.Count - 1];
            fragments[fragments.Count - 1] = new PositionedTextFragment(previous.Text + text, style);
            return;
        }
        fragments.Add(new PositionedTextFragment(text, style));
    }

    private static bool IsSafelyEditableSpan(PdfTextSpan span, bool allowTextRenderingMode3 = false) =>
        ((span.IsVisible && span.TextRenderingMode == 0) ||
         (allowTextRenderingMode3 && !span.IsVisible && span.TextRenderingMode == 3)) &&
        !(span.TextRenderingMode == 3 && span.IsType3Font) &&
        !span.ClipPath.HasValue &&
        (span.TextRenderingMode == 3 || !span.Color.HasValue || span.Color.Value.A == byte.MaxValue) &&
        span.CanRestamp &&
        !string.IsNullOrEmpty(span.Text);

    private static bool IsSearchSpan(PdfTextSpan span, bool includeTextRenderingMode3) =>
        (span.IsVisible || (includeTextRenderingMode3 && span.TextRenderingMode == 3 && !span.IsType3Font)) &&
        !span.ClipPath.HasValue &&
        (span.TextRenderingMode == 3 || !span.Color.HasValue || span.Color.Value.A > 0) &&
        !string.IsNullOrEmpty(span.Text);

    private static PdfReadDocument OpenForVisualTextEditing(byte[] pdf, PdfLoadOptions? readOptions) =>
        PdfReadDocument.Open(pdf, PdfLoadOptions.WithArtifactText(readOptions));

    private static PdfPageRegion TranslateRegionToSource(byte[] pdf, PdfPageRegion region, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(region, nameof(region));
        PdfReadDocument document = OpenForVisualTextEditing(pdf, readOptions);
        ValidatePage(region.PageNumber, document.Pages.Count, nameof(region));
        (double originX, double originY) = document.Pages[region.PageNumber - 1].GetPageBoundaryOrigin();
        return OffsetRegion(region, originX, originY);
    }

    private static PdfPageRegion OffsetRegion(PdfPageRegion region, double deltaX, double deltaY) =>
        new PdfPageRegion(region.PageNumber, region.X + deltaX, region.Y + deltaY, region.Width, region.Height);

    private static PdfRegionText InspectSource(byte[] pdf, PdfPageRegion region, PdfLoadOptions? readOptions) =>
        InspectSource(pdf, region, includeTextRenderingMode3: false, readOptions);

    private static PdfRegionText InspectSource(byte[] pdf, PdfPageRegion region, bool includeTextRenderingMode3, PdfLoadOptions? readOptions) {
        PdfReadDocument document = OpenForVisualTextEditing(pdf, readOptions);
        ValidatePage(region.PageNumber, document.Pages.Count, nameof(region));
        return InspectSource(document.Pages[region.PageNumber - 1], region, includeTextRenderingMode3);
    }

    private static PdfRegionText InspectSource(PdfReadPage page, PdfPageRegion region, bool includeTextRenderingMode3 = false) =>
        BuildRegionText(page.GetTextSpans()
            .Where(span => IsSearchSpan(span, includeTextRenderingMode3) && Intersects(region, GetBounds(span)))
            .ToArray());

    private static void EnsureRegionIsSafelyEditable(byte[] pdf, PdfPageRegion region, bool allowTextRenderingMode3, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(region, nameof(region));
        PdfReadDocument document = OpenForVisualTextEditing(pdf, readOptions);
        ValidatePage(region.PageNumber, document.Pages.Count, nameof(region));
        bool containsUnsafeText = document.Pages[region.PageNumber - 1]
            .GetTextSpans()
            .Any(span => !IsSafelyEditableSpan(span, allowTextRenderingMode3) && Intersects(region, GetBounds(span)));
        if (containsUnsafeText) {
            throw new NotSupportedException("The selected region contains invisible or clipped text whose rendering state cannot be recreated safely.");
        }
    }

    private static void EnsureCompatibleRenderingModes(IReadOnlyList<PdfTextSpan> spans, bool allowTextRenderingMode3) {
        for (int index = 0; index < spans.Count; index++) {
            if (!IsSafelyEditableSpan(spans[index], allowTextRenderingMode3)) {
                throw new NotSupportedException("The selected region contains invisible or clipped text whose rendering state cannot be recreated safely.");
            }
        }
        if (spans.Any(static span => span.TextRenderingMode == 3) &&
            spans.Any(static span => span.TextRenderingMode != 3)) {
            throw new NotSupportedException("One text edit cannot combine visible text with rendering-mode-3 OCR text.");
        }
    }

    private static void EnsureAppendOrderIsSafe(byte[] pdf, int pageNumber, IReadOnlyList<PdfTextSpan> spans, PdfLoadOptions? readOptions, IReadOnlyList<PdfStamper.TextStampRequest>? appendedRequests = null) {
        if (spans.Count == 0) return;
        PdfReadDocument document = OpenForVisualTextEditing(pdf, readOptions);
        PdfTextSpan[] visibleSpans = spans.Where(static span => span.TextRenderingMode != 3).ToArray();
        PdfReadPage.PdfAppendedTextBounds[]? appendedBounds = appendedRequests?
            .Where(static request => request.TextRenderingMode != 3)
            .Select(ToAppendedTextBounds)
            .ToArray();
        if (visibleSpans.Length == 0 && (appendedBounds == null || appendedBounds.Length == 0)) return;
        if (document.Pages[pageNumber - 1].WouldAppendingTextChangeVisibleStacking(visibleSpans, appendedBounds)) {
            throw new NotSupportedException("The text edit would change the visible stacking order of overlapping page content.");
        }
    }

    private static PdfReadPage.PdfAppendedTextBounds ToAppendedTextBounds(PdfStamper.TextStampRequest request) {
        double advance = Math.Max(0.1D, PdfWriter.EstimateSimpleTextWidth(request.Text, request.Font, request.FontSize));
        double radians = request.RotationDegrees * Math.PI / 180D;
        double ux = Math.Cos(radians);
        double uy = Math.Sin(radians);
        double nx = -uy;
        double ny = ux;
        double descent = request.FontSize * 0.25D;
        double ascent = request.FontSize * 0.8D;
        double[] x = {
            request.X - (nx * descent),
            request.X + (ux * advance) - (nx * descent),
            request.X + (nx * ascent),
            request.X + (ux * advance) + (nx * ascent)
        };
        double[] y = {
            request.Y - (ny * descent),
            request.Y + (uy * advance) - (ny * descent),
            request.Y + (ny * ascent),
            request.Y + (uy * advance) + (ny * ascent)
        };
        return new PdfReadPage.PdfAppendedTextBounds(x.Min(), y.Min(), x.Max(), y.Max(), request.PaintOrder);
    }

    private static void ValidatePage(int pageNumber, int pageCount, string paramName) {
        if (pageNumber > pageCount) throw new ArgumentOutOfRangeException(paramName, "Page number exceeds the PDF page count.");
    }

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(paramName, "Offset must be finite.");
    }

    internal sealed class TextMutationResult {
        internal TextMutationResult(byte[] bytes, int affectedCount, IEnumerable<string> warnings) {
            Bytes = bytes;
            AffectedCount = affectedCount;
            Warnings = warnings.Distinct(StringComparer.Ordinal).ToArray();
        }
        internal byte[] Bytes { get; }
        internal int AffectedCount { get; }
        internal IReadOnlyList<string> Warnings { get; }
    }

    private readonly struct PdfResolvedTextStyle : IEquatable<PdfResolvedTextStyle> {
        internal PdfResolvedTextStyle(PdfStandardFont font, double fontSize, PdfColor color, double rotationDegrees, int textRenderingMode) { Font = font; FontSize = fontSize; Color = color; RotationDegrees = rotationDegrees; TextRenderingMode = textRenderingMode; }
        internal PdfStandardFont Font { get; }
        internal double FontSize { get; }
        internal PdfColor Color { get; }
        internal double RotationDegrees { get; }
        internal int TextRenderingMode { get; }
        public bool Equals(PdfResolvedTextStyle other) => Font == other.Font && FontSize.Equals(other.FontSize) && Color.Equals(other.Color) && RotationDegrees.Equals(other.RotationDegrees) && TextRenderingMode == other.TextRenderingMode;
        public override bool Equals(object? obj) => obj is PdfResolvedTextStyle other && Equals(other);
        public override int GetHashCode() { unchecked { int hash = (int)Font; hash = (hash * 397) ^ FontSize.GetHashCode(); hash = (hash * 397) ^ Color.GetHashCode(); hash = (hash * 397) ^ RotationDegrees.GetHashCode(); return (hash * 397) ^ TextRenderingMode; } }
    }

    private readonly struct SpanBounds {
        internal SpanBounds(double x, double y, double width, double height) { X = x; Y = y; Width = width; Height = height; }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
    }

    private readonly struct SpanStyleKey : IEquatable<SpanStyleKey> {
        internal SpanStyleKey(string font, double size, double rotation, OfficeColor? color) { Font = font; Size = size; Rotation = rotation; Color = color; }
        private string Font { get; }
        private double Size { get; }
        private double Rotation { get; }
        private OfficeColor? Color { get; }
        public bool Equals(SpanStyleKey other) => string.Equals(Font, other.Font, StringComparison.Ordinal) && Size.Equals(other.Size) && Rotation.Equals(other.Rotation) && Nullable.Equals(Color, other.Color);
        public override bool Equals(object? obj) => obj is SpanStyleKey other && Equals(other);
        public override int GetHashCode() { unchecked { int hash = StringComparer.Ordinal.GetHashCode(Font); hash = (hash * 397) ^ Size.GetHashCode(); hash = (hash * 397) ^ Rotation.GetHashCode(); return (hash * 397) ^ Color.GetHashCode(); } }
    }

    private readonly struct PageTextSpanSnapshot {
        internal PageTextSpanSnapshot(int pageNumber, PdfTextSpan span, bool targeted) { PageNumber = pageNumber; Span = span; Targeted = targeted; }
        internal int PageNumber { get; }
        internal PdfTextSpan Span { get; }
        internal bool Targeted { get; }
    }

    private readonly struct TextRemovalResult {
        internal TextRemovalResult(byte[] bytes, IEnumerable<string> warnings, IEnumerable<PdfStamper.TextStampRequest> restamps) { Bytes = bytes; Warnings = warnings.Distinct(StringComparer.Ordinal).ToArray(); Restamps = restamps.ToArray(); }
        internal byte[] Bytes { get; }
        internal IReadOnlyList<string> Warnings { get; }
        internal IReadOnlyList<PdfStamper.TextStampRequest> Restamps { get; }
    }

    private readonly struct PositionedRewrite {
        internal PositionedRewrite(int pageNumber, PdfTextSpan source, PositionedTextFragment[] fragments) { PageNumber = pageNumber; Source = source; Fragments = fragments; }
        internal int PageNumber { get; }
        internal PdfTextSpan Source { get; }
        internal PositionedTextFragment[] Fragments { get; }
    }

    private readonly struct PositionedTextFragment {
        internal PositionedTextFragment(string text, PdfResolvedTextStyle style) { Text = text; Style = style; }
        internal string Text { get; }
        internal PdfResolvedTextStyle Style { get; }
    }

    private readonly struct RewriteLineKey : IEquatable<RewriteLineKey> {
        private RewriteLineKey(int pageNumber, double rotation, bool isTextRenderingMode3) { PageNumber = pageNumber; Rotation = rotation; IsTextRenderingMode3 = isTextRenderingMode3; }
        private int PageNumber { get; }
        private double Rotation { get; }
        private bool IsTextRenderingMode3 { get; }
        internal static RewriteLineKey Create(PositionedRewrite rewrite) => new RewriteLineKey(rewrite.PageNumber, Math.Round(rewrite.Source.RotationDegrees, 1), rewrite.Source.TextRenderingMode == 3);
        public bool Equals(RewriteLineKey other) => PageNumber == other.PageNumber && Rotation.Equals(other.Rotation) && IsTextRenderingMode3 == other.IsTextRenderingMode3;
        public override bool Equals(object? obj) => obj is RewriteLineKey other && Equals(other);
        public override int GetHashCode() { unchecked { return ((PageNumber * 397) ^ Rotation.GetHashCode()) * 397 ^ IsTextRenderingMode3.GetHashCode(); } }
    }

    private readonly struct SpanTextEdit {
        internal SpanTextEdit(int start, int length, string replacement) { Start = start; Length = length; Replacement = replacement; }
        internal int Start { get; }
        internal int Length { get; }
        internal string Replacement { get; }
    }

    private readonly struct PageSpanKey : IEquatable<PageSpanKey> {
        internal PageSpanKey(int pageNumber, PdfTextSpan span) { PageNumber = pageNumber; Span = span; }
        internal int PageNumber { get; }
        internal PdfTextSpan Span { get; }
        public bool Equals(PageSpanKey other) => PageNumber == other.PageNumber && ReferenceEquals(Span, other.Span);
        public override bool Equals(object? obj) => obj is PageSpanKey other && Equals(other);
        public override int GetHashCode() { unchecked { return (PageNumber * 397) ^ System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(Span); } }
    }

    private sealed class TextSpanArrayReferenceComparer : IEqualityComparer<PdfTextSpan[]> {
        internal static TextSpanArrayReferenceComparer Instance { get; } = new TextSpanArrayReferenceComparer();

        public bool Equals(PdfTextSpan[]? x, PdfTextSpan[]? y) => ReferenceEquals(x, y);

        public int GetHashCode(PdfTextSpan[] value) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
    }
}

using System.Text.RegularExpressions;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfTextEditor {
    // A search mark may reach into a neighbouring line, but a kept glyph must stay readable. Past this
    // share of a glyph's body the mark hides the glyph, or the text overlaps the match: either way the
    // selection cannot be treated as separate from it.
    private const double MaximumNeighbourBodyCoverage = 0.25D;

    internal static bool HasRegexRedactionMatch(byte[] pdf, Regex expression, int pageNumber, PdfLoadOptions? readOptions,
        PdfRedactionSearchWorkBudget budget, CancellationToken cancellationToken) =>
        FindHits(pdf, expression.ToString(), new PdfTextSearchOptions { IncludeTextRenderingMode3 = true, PageNumbers = new[] { pageNumber } },
            readOptions, budget, unit => budget.Matches(expression, unit.Text)
                .Where(static match => match.Length > 0).Select(static match => new TextSearchRange(match.Index, match.Length, match.Value)),
            includeUnsafeSpans: true, cancellationToken: cancellationToken).Count > 0;

    // Hidden layers are read separately so their text cannot interleave with the visible lines.
    internal static bool HasHiddenLayerRedactionMatch(byte[] pdf, string literal, bool matchCase, int pageNumber, PdfLoadOptions? readOptions,
        PdfRedactionSearchWorkBudget budget, CancellationToken cancellationToken) =>
        HasOptionalContent(pdf, readOptions, cancellationToken) &&
        FindHits(pdf, literal, new PdfTextSearchOptions { MatchCase = matchCase, IncludeTextRenderingMode3 = true, PageNumbers = new[] { pageNumber } },
            readOptions, budget, includeUnsafeSpans: true, hiddenOptionalContentOnly: true, cancellationToken: cancellationToken).Count > 0;

    internal static bool HasHiddenLayerRegexRedactionMatch(byte[] pdf, Regex expression, int pageNumber, PdfLoadOptions? readOptions,
        PdfRedactionSearchWorkBudget budget, CancellationToken cancellationToken) =>
        HasOptionalContent(pdf, readOptions, cancellationToken) &&
        FindHits(pdf, expression.ToString(), new PdfTextSearchOptions { IncludeTextRenderingMode3 = true, PageNumbers = new[] { pageNumber } },
            readOptions, budget, unit => budget.Matches(expression, unit.Text)
                .Where(static match => match.Length > 0).Select(static match => new TextSearchRange(match.Index, match.Length, match.Value)),
            includeUnsafeSpans: true, hiddenOptionalContentOnly: true, cancellationToken: cancellationToken).Count > 0;

    private static bool HasOptionalContent(byte[] pdf, PdfLoadOptions? readOptions, CancellationToken cancellationToken) =>
        PdfReadDocument.Open(pdf, readOptions, cancellationToken).CatalogDictionary?.Items.ContainsKey("OCProperties") == true;

    /// <summary>The matched lines of a hit, for callers that select complete lines rather than glyph ranges.</summary>
    internal static IReadOnlyList<(PdfTextMatch Match, PdfTextSpan[][] Lines)> FindMatchedLines(byte[] pdf, string literal, PdfTextSearchOptions options,
        PdfLoadOptions? readOptions, PdfRedactionSearchWorkBudget budget, CancellationToken cancellationToken) =>
        FindHits(pdf, literal, options, readOptions, budget, cancellationToken: cancellationToken)
            .Select(static hit => (hit.Match, hit.Lines)).ToArray();

    // Logical blocks don't include hidden layers, so a hidden match selects the whole hidden line that holds it.
    internal static IReadOnlyList<PdfRedactionBlockCandidate> FindHiddenLayerBlockCandidates(
        byte[] pdf, PdfRedactionSearchOptions search, Regex[] expressions, PdfLoadOptions? readOptions, PdfRedactionSearchWorkBudget budget) {
        var candidates = new List<PdfRedactionBlockCandidate>();
        if (!HasOptionalContent(pdf, readOptions, search.CancellationToken)) return candidates;
        var options = new PdfTextSearchOptions {
            MatchCase = search.MatchCase, IncludeTextRenderingMode3 = true, PageNumbers = search.PageNumbers.ToArray()
        };
        foreach (string literal in search.LiteralText) {
            AddLines(FindHits(pdf, literal, options, readOptions, budget, includeUnsafeSpans: true, hiddenOptionalContentOnly: true,
                cancellationToken: search.CancellationToken, maximumHits: search.MaximumCandidates), "literal:" + literal);
        }
        for (int index = 0; index < expressions.Length; index++) {
            Regex expression = expressions[index];
            AddLines(FindHits(pdf, expression.ToString(), options, readOptions, budget,
                unit => budget.Matches(expression, unit.Text).Where(static match => match.Length > 0)
                    .Select(static match => new TextSearchRange(match.Index, match.Length, match.Value)),
                includeUnsafeSpans: true, hiddenOptionalContentOnly: true,
                cancellationToken: search.CancellationToken, maximumHits: search.MaximumCandidates), "regex:" + search.RegularExpressions[index]);
        }
        return candidates;

        void AddLines(IReadOnlyList<TextSearchHit> hits, string label) {
            foreach (TextSearchHit hit in hits) {
                foreach (PdfTextSpan[] line in hit.Lines) {
                    PdfTextSpanBounds[] bounds = line.Select(PdfTextSpanGeometry.GetRedactionBounds).ToArray();
                    candidates.Add(new PdfRedactionBlockCandidate(hit.PageNumber, new PdfTextSpanBounds(bounds.Min(static value => value.Left), bounds.Min(static value => value.Bottom),
                        bounds.Max(static value => value.Right), bounds.Max(static value => value.Top)), label, line, isHiddenOnly: true));
                }
            }
        }
    }

    /// <summary>
    /// Turns complete-block candidates into glyph selections, so a block takes only its own glyphs while its
    /// mark keeps the full block rectangle. Anything that can't be tied to a block is refused like a precise match.
    /// </summary>
    internal static IReadOnlyList<PdfRedactionArea> SelectBlockGlyphs(byte[] pdf, IReadOnlyList<PdfRedactionBlockCandidate> candidates,
        PdfRedactionSearchOptions search, PdfLoadOptions? readOptions, PdfRedactionSearchWorkBudget budget, out IReadOnlyList<PdfDiagnosticFinding> findings) {
        var diagnostics = new List<PdfDiagnosticFinding>();
        var diagnosticKeys = new HashSet<(int Page, string Code, string Message)>();
        void Block(int page, string code, string message) {
            if (diagnosticKeys.Add((page, code, message))) diagnostics.Add(new PdfDiagnosticFinding(PdfDiagnosticSeverity.Error, code, message, pageNumber: page));
        }
        var areas = new List<PdfRedactionArea>();
        var rectangleAreas = new List<PdfRedactionArea>();
        var selected = new HashSet<(int Page, PdfContentOrderKey Owner, int Run, int TextStart)>();
        PdfReadDocument document = PdfReadDocument.Open(pdf, PdfLoadOptions.WithArtifactText(readOptions), search.CancellationToken);
        foreach (IGrouping<int, PdfRedactionBlockCandidate> page in candidates.GroupBy(static candidate => candidate.PageNumber)) {
            search.CancellationToken.ThrowIfCancellationRequested();
            PdfReadPage readPage = document.Pages[page.Key - 1];
            IReadOnlyList<PdfTextSpan> spans = readPage.GetGlyphTextSpans(includeHiddenOptionalContent: true, cancellationToken: search.CancellationToken);
            var visibleKeys = new HashSet<PdfContentOrderKey>(readPage.GetGlyphTextSpans(cancellationToken: search.CancellationToken)
                .Select(static span => span.ContentOrderKey).OfType<PdfContentOrderKey>());
            foreach (PdfRedactionBlockCandidate candidate in page) {
                // A block belongs to one layer: overlapping text in the other layer is not part of it.
                PdfTextSpan[] layer = spans.Where(span => candidate.IsHiddenOnly ==
                    !(span.ContentOrderKey is PdfContentOrderKey key && visibleKeys.Contains(key))).ToArray();
                bool Touches(PdfTextSpan span) {
                    PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetRedactionBounds(span);
                    return bounds.Right > candidate.Bounds.Left && bounds.Left < candidate.Bounds.Right &&
                        bounds.Top > candidate.Bounds.Bottom && bounds.Bottom < candidate.Bounds.Top;
                }
                PdfTextSpan[] unmapped = layer.Where(span => !string.IsNullOrWhiteSpace(span.Text) && !CanSelectGlyphs(span, out _) && Touches(span) &&
                    candidate.Owners.Any(owner => IsOnOwnerBaseline(owner, span.X, span.Y))).ToArray();
                if (unmapped.Length > 0) {
                    // Text with no glyph mapping can only go as a rectangle, which is safe when nothing else is in reach.
                    if (TryAddUnmappedBlock(candidate, unmapped, spans)) continue;
                    Block(page.Key, "redaction-search-glyph-mapping-unsupported", GlyphMappingMessage);
                    continue;
                }
                var glyphs = new List<PdfRedactionSelectedGlyph>();
                foreach (PdfTextSpan span in layer) {
                    if (!CanSelectGlyphs(span, out IReadOnlyList<PdfTextGlyph> spanGlyphs)) continue;
                    PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries);
                    foreach (PdfTextGlyph glyph in spanGlyphs) {
                        budget.Charge(1L);
                        PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetRedactionGlyphBounds(span, boundaries[glyph.TextStart], glyph.PaintedAdvance);
                        if (bounds.Right <= candidate.Bounds.Left || bounds.Left >= candidate.Bounds.Right ||
                            bounds.Top <= candidate.Bounds.Bottom || bounds.Bottom >= candidate.Bounds.Top ||
                            !candidate.Owners.Any(owner => IsOnOwnerBaseline(owner, glyph.X, glyph.Y))) continue;
                        selected.Add(GlyphKey(page.Key, span, glyph));
                        glyphs.Add(new PdfRedactionSelectedGlyph(glyph.X, glyph.Y, glyph.EncodedBytes));
                    }
                }
                if (glyphs.Count == 0) {
                    Block(page.Key, "redaction-search-glyph-mapping-unsupported", GlyphMappingMessage);
                    continue;
                }
                PdfRedactionContentScope scope = candidate.IsHiddenOnly ? PdfRedactionContentScope.TextOnly : search.ContentScope;
                areas.Add(new PdfRedactionArea(page.Key, candidate.Bounds.Left, candidate.Bounds.Bottom, candidate.Bounds.Right - candidate.Bounds.Left,
                    candidate.Bounds.Top - candidate.Bounds.Bottom, candidate.Label, scope)
                    .WithGlyphSelection(new PdfRedactionGlyphSelection(glyphs, candidate.IsHiddenOnly, isCompleteBlock: true)));
            }

            bool TryAddUnmappedBlock(PdfRedactionBlockCandidate candidate, PdfTextSpan[] unmapped, IReadOnlyList<PdfTextSpan> allSpans) {
                var objectKeys = new HashSet<PdfContentOrderKey>(unmapped.Select(static span => span.TextObjectOrderKey).OfType<PdfContentOrderKey>());
                foreach (PdfTextSpan span in allSpans) {
                    if (unmapped.Contains(span) || string.IsNullOrWhiteSpace(span.Text) && !(span.TextObjectOrderKey is PdfContentOrderKey blank && objectKeys.Contains(blank))) continue;
                    PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetRedactionBounds(span);
                    bool inReach = bounds.Right > candidate.Bounds.Left && bounds.Left < candidate.Bounds.Right &&
                        bounds.Top > candidate.Bounds.Bottom && bounds.Bottom < candidate.Bounds.Top;
                    // A failed glyph rewrite falls back to the whole text object, so nothing else may share it.
                    if (inReach || span.TextObjectOrderKey is PdfContentOrderKey key && objectKeys.Contains(key)) return false;
                }
                PdfRedactionContentScope scope = candidate.IsHiddenOnly ? PdfRedactionContentScope.TextOnly : search.ContentScope;
                rectangleAreas.Add(new PdfRedactionArea(candidate.PageNumber, candidate.Bounds.Left, candidate.Bounds.Bottom, candidate.Bounds.Right - candidate.Bounds.Left,
                    candidate.Bounds.Top - candidate.Bounds.Bottom, candidate.Label, scope));
                return true;
            }
        }
        CheckSelectionSeparation(document, areas, selected, budget, Block, search.CancellationToken);
        areas.AddRange(rectangleAreas);
        findings = diagnostics.AsReadOnly();
        return areas.AsReadOnly();
    }

    private static bool IsOnOwnerBaseline(PdfTextSpan owner, double x, double y) {
        const double tolerance = 0.01D;
        double radians = owner.RotationDegrees * Math.PI / 180D;
        double along = Math.Cos(radians), up = Math.Sin(radians);
        double deltaX = x - owner.X, deltaY = y - owner.Y;
        double position = deltaX * along + deltaY * up;
        return Math.Abs(deltaY * along - deltaX * up) <= tolerance && position >= -tolerance && position <= Math.Abs(owner.Advance) + tolerance;
    }

    private static void CheckSelectionSeparation(PdfReadDocument document, IReadOnlyList<PdfRedactionArea> areas,
        HashSet<(int Page, PdfContentOrderKey Owner, int Run, int TextStart)> selected, PdfRedactionSearchWorkBudget budget,
        Action<int, string, string> block, CancellationToken cancellationToken) {
        foreach (IGrouping<int, PdfRedactionArea> pageAreas in areas.GroupBy(static area => area.PageNumber)) {
            PdfReadPage page = document.Pages[pageAreas.Key - 1];
            // A hidden-only area paints nothing, so it can't cover anything; it only has to stay unambiguous.
            PdfRedactionArea[] markedAreas = pageAreas.Where(static area => !area.IsHiddenOnly).ToArray();
            PdfRedactionGlyphSelection[] hiddenSelections = pageAreas.Where(static area => area.IsHiddenOnly).Select(static area => area.GlyphSelection!).ToArray();
            // TextOnly keeps the image under the mark, so a nearby line would be left over retained pixels
            // of the match. Those areas keep the strict rectangle check; full removal wipes the pixels.
            IReadOnlyList<PdfImagePlacement> images = page.GetImagePlacementsIncludingHiddenOptionalContent(pageAreas.Key);
            PdfRedactionArea[] imageAreas = markedAreas.Where(area => area.ContentScope == PdfRedactionContentScope.TextOnly &&
                images.Any(image => area.IntersectsRectangle(image.X, image.Y, image.Width, image.Height))).ToArray();
            foreach (PdfTextSpan span in page.GetGlyphTextSpans(includeHiddenOptionalContent: true, cancellationToken: cancellationToken)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!CanSelectGlyphs(span, out IReadOnlyList<PdfTextGlyph> glyphs)) {
                    if (string.IsNullOrWhiteSpace(span.Text)) continue;
                    if (IntersectsAreas(markedAreas, PdfTextSpanGeometry.GetGlyphBodyBounds(span, 0D, span.Advance)) ||
                        IntersectsAreas(imageAreas, PdfTextSpanGeometry.GetRedactionGlyphBounds(span, 0D, span.Advance)))
                        block(pageAreas.Key, "redaction-search-glyph-mapping-unsupported", GlyphMappingMessage);
                    continue;
                }
                PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries);
                foreach (PdfTextGlyph glyph in glyphs) {
                    budget.Charge(1L);
                    if (selected.Contains(GlyphKey(pageAreas.Key, span, glyph))) continue;
                    double offset = boundaries[glyph.TextStart];
                    if (IntersectsAreas(imageAreas, PdfTextSpanGeometry.GetRedactionGlyphBounds(span, offset, glyph.PaintedAdvance)) ||
                        hiddenSelections.Any(selection => selection.Contains(glyph.X, glyph.Y, glyph.EncodedBytes)))
                        block(pageAreas.Key, "redaction-search-unselected-text-intersection", UnselectedTextMessage);
                    if (string.IsNullOrWhiteSpace(glyph.Text)) continue;
                    if (GetCoveredShare(markedAreas, PdfTextSpanGeometry.GetGlyphBodyBounds(span, offset, glyph.PaintedAdvance)) > MaximumNeighbourBodyCoverage)
                        block(pageAreas.Key, "redaction-search-unselected-text-intersection", UnselectedTextMessage);
                }
            }
        }

        bool IntersectsAreas(IEnumerable<PdfRedactionArea> candidates, PdfTextSpanBounds bounds) {
            foreach (PdfRedactionArea area in candidates) {
                budget.Charge(1L);
                if (area.IntersectsRectangle(bounds.Left, bounds.Bottom, bounds.Width, bounds.Height)) return true;
            }
            return false;
        }

        double GetCoveredShare(IEnumerable<PdfRedactionArea> candidates, PdfTextSpanBounds body) {
            double covered = 0D;
            double total = (body.Right - body.Left) * (body.Top - body.Bottom);
            if (total <= 0D) return 0D;
            foreach (PdfRedactionArea area in candidates) {
                budget.Charge(1L);
                double width = Math.Min(area.Right, body.Right) - Math.Max(area.X, body.Left);
                double height = Math.Min(area.Top, body.Top) - Math.Max(area.Y, body.Bottom);
                if (width > 0D && height > 0D) covered += width * height;
            }
            return covered / total;
        }
    }

    // Reuse the editor's source-range, line-flow and table boundaries instead of
    // approximating a substring's position from logical-block text lengths.
    internal static IReadOnlyList<PdfRedactionArea> FindRedactionAreas(
        byte[] pdf, PdfRedactionSearchOptions search, Regex[] expressions, PdfLoadOptions? readOptions,
        PdfRedactionSearchWorkBudget budget, out IReadOnlyList<PdfDiagnosticFinding> findings) {
        var drafts = new List<MatchedGlyphDraft>();
        var draftKeys = new Dictionary<(int Page, double Left, double Bottom, double Right, double Top), MatchedGlyphDraft>();
        var diagnostics = new List<PdfDiagnosticFinding>();
        var selected = new HashSet<(int Page, PdfContentOrderKey Owner, int Run, int TextStart)>();
        var diagnosticKeys = new HashSet<(int Page, string Code, string Message)>();
        var options = new PdfTextSearchOptions {
            MatchCase = search.MatchCase, IncludeTextRenderingMode3 = true,
            PageNumbers = search.PageNumbers.ToArray()
        };
        bool searchHiddenLayers = HasOptionalContent(pdf, readOptions, search.CancellationToken);
        for (int index = 0; index < search.LiteralText.Count; index++) {
            string literal = search.LiteralText[index];
            AddHits(FindHits(pdf, literal, options, readOptions, budget, includeUnsafeSpans: true,
                cancellationToken: search.CancellationToken, maximumHits: search.MaximumCandidates), "literal:" + literal, false);
            if (searchHiddenLayers) AddHits(FindHits(pdf, literal, options, readOptions, budget, includeUnsafeSpans: true, hiddenOptionalContentOnly: true,
                cancellationToken: search.CancellationToken, maximumHits: search.MaximumCandidates), "literal:" + literal, true);
        }
        for (int index = 0; index < expressions.Length; index++) {
            Regex expression = expressions[index];
            AddHits(FindHits(pdf, expression.ToString(), options, readOptions, budget,
                unit => RegexRanges(unit.Text, expression), includeUnsafeSpans: true,
                cancellationToken: search.CancellationToken, maximumHits: search.MaximumCandidates),
                "regex:" + search.RegularExpressions[index], false);
            if (searchHiddenLayers) AddHits(FindHits(pdf, expression.ToString(), options, readOptions, budget,
                unit => RegexRanges(unit.Text, expression), includeUnsafeSpans: true, hiddenOptionalContentOnly: true,
                cancellationToken: search.CancellationToken, maximumHits: search.MaximumCandidates),
                "regex:" + search.RegularExpressions[index], true);
        }
        PdfRedactionArea[] areas = drafts
            .GroupBy(static draft => (draft.Page, draft.Left, draft.Bottom, draft.Right, draft.Top))
            .Select(group => {
                MatchedGlyphDraft first = group.First();
                // A rectangle shared with a visible match still needs its mark.
                bool hiddenOnly = group.All(static draft => draft.IsHiddenOnly);
                return new PdfRedactionArea(first.Page, first.Left, first.Bottom, first.Right - first.Left, first.Top - first.Bottom, first.Label,
                    hiddenOnly ? PdfRedactionContentScope.TextOnly : search.ContentScope)
                    .WithGlyphSelection(new PdfRedactionGlyphSelection(group.SelectMany(static draft => draft.Glyphs), hiddenOnly));
            }).ToArray();
        if (areas.Length > 0) {
            PdfReadDocument document = PdfReadDocument.Open(pdf, PdfLoadOptions.WithArtifactText(readOptions), search.CancellationToken);
            CheckSelectionSeparation(document, areas, selected, budget, Block, search.CancellationToken);
        }
        findings = diagnostics.AsReadOnly();
        return areas;

        IEnumerable<TextSearchRange> RegexRanges(string text, Regex expression) {
            foreach (Match match in budget.Matches(expression, text)) {
                search.CancellationToken.ThrowIfCancellationRequested();
                if (match.Length == 0) throw new InvalidOperationException("A precise redaction regex must select non-empty source text.");
                yield return new TextSearchRange(match.Index, match.Length, match.Value);
            }
        }

        void AddHits(IReadOnlyList<TextSearchHit> hits, string label, bool hiddenOnly) {
            foreach (TextSearchHit hit in hits) {
                int firstDraft = drafts.Count;
                PdfTextSpan? previousSpan = null;
                foreach (TextSourceSegment segment in hit.Segments) {
                    search.CancellationToken.ThrowIfCancellationRequested();
                    PdfTextSpan span = segment.Span;
                    if (!CanSelectGlyphs(span, out IReadOnlyList<PdfTextGlyph> glyphs)) {
                        Block(hit.PageNumber, "redaction-search-glyph-mapping-unsupported", GlyphMappingMessage);
                        continue;
                    }
                    int end = segment.Start + segment.Length;
                    PdfTextGlyph[] covered = glyphs.Where(glyph => glyph.TextStart >= segment.Start && glyph.TextStart + glyph.TextLength <= end).ToArray();
                    if (covered.Length == 0 || covered[0].TextStart != segment.Start ||
                        covered[covered.Length - 1].TextStart + covered[covered.Length - 1].TextLength != end) {
                        Block(hit.PageNumber, "redaction-search-partial-glyph",
                            "The match covers only part of a joined character, such as a ligature. Search for the whole word or draw the area by hand.");
                        continue;
                    }
                    PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries);
                    double left = double.PositiveInfinity, bottom = double.PositiveInfinity;
                    double right = double.NegativeInfinity, top = double.NegativeInfinity;
                    var coveredGlyphs = new List<PdfRedactionSelectedGlyph>(covered.Length);
                    foreach (PdfTextGlyph glyph in covered) {
                        budget.Charge(1L);
                        selected.Add(GlyphKey(hit.PageNumber, span, glyph));
                        coveredGlyphs.Add(new PdfRedactionSelectedGlyph(glyph.X, glyph.Y, glyph.EncodedBytes));
                        PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetRedactionGlyphBounds(span, boundaries[glyph.TextStart], glyph.PaintedAdvance);
                        left = Math.Min(left, bounds.Left); bottom = Math.Min(bottom, bounds.Bottom);
                        right = Math.Max(right, bounds.Right); top = Math.Max(top, bounds.Top);
                    }
                    (int, double, double, double, double) key = (hit.PageNumber, left, bottom, right, top);
                    if (draftKeys.TryGetValue(key, out MatchedGlyphDraft? existing)) {
                        // Coincident glyphs share one rectangle, and each of them still has to go.
                        existing.Glyphs.AddRange(coveredGlyphs);
                        continue;
                    }
                    // A phrase may cross adjacent native text shows (including explicit spaces).
                    // Small kerning gaps inside one TJ operator belong to the same phrase geometry.
                    // Independent text operators still have to touch; wide TJ positioning remains separate.
                    double joinTolerance = previousSpan?.ContentOrderKey is not null && previousSpan.ContentOrderKey.Equals(span.ContentOrderKey)
                        ? Math.Max(0.0000001D, Math.Min(previousSpan.FontSize, span.FontSize) * 0.25D) : 0.0000001D;
                    if (drafts.Count > firstDraft && drafts[drafts.Count - 1].CanJoin(hit.PageNumber, left, bottom, right, top, joinTolerance)) {
                        MatchedGlyphDraft previous = drafts[drafts.Count - 1];
                        previous.Join(left, bottom, right, top);
                        previous.Glyphs.AddRange(coveredGlyphs);
                        draftKeys[key] = previous;
                        previousSpan = span;
                        continue;
                    }
                    if (drafts.Count >= search.MaximumCandidates) throw new InvalidOperationException("Redaction search exceeded the configured candidate limit.");
                    var draft = new MatchedGlyphDraft(hit.PageNumber, left, bottom, right, top, label, hiddenOnly);
                    draft.Glyphs.AddRange(coveredGlyphs);
                    drafts.Add(draft);
                    draftKeys[key] = draft;
                    previousSpan = span;
                }
            }
        }

        void Block(int page, string code, string message) {
            if (diagnosticKeys.Add((page, code, message))) diagnostics.Add(new PdfDiagnosticFinding(PdfDiagnosticSeverity.Error, code, message, pageNumber: page));
        }
    }

    // Studio matches on this text, so keep the wording as it is.
    private const string UnselectedTextMessage =
        "A precise search area would remove an unselected glyph. Use reviewed geometry or complete-block selection instead.";

    private const string GlyphMappingMessage =
        "This text can't be traced to exact characters in the file, so it can't be removed on its own. Select the whole block or draw the area by hand.";

    private sealed class MatchedGlyphDraft {
        internal MatchedGlyphDraft(int page, double left, double bottom, double right, double top, string label, bool isHiddenOnly) {
            Page = page; Left = left; Bottom = bottom; Right = right; Top = top; Label = label; IsHiddenOnly = isHiddenOnly;
        }

        internal bool IsHiddenOnly { get; }

        internal int Page { get; }
        internal double Left { get; private set; }
        internal double Bottom { get; private set; }
        internal double Right { get; private set; }
        internal double Top { get; private set; }
        internal string Label { get; }
        internal List<PdfRedactionSelectedGlyph> Glyphs { get; } = new();

        internal bool CanJoin(int page, double left, double bottom, double right, double top, double tolerance) =>
            Page == page && (
                Math.Abs(Bottom - bottom) < 0.0000001D && Math.Abs(Top - top) < 0.0000001D &&
                left <= Right + tolerance && right >= Left - tolerance ||
                Math.Abs(Left - left) < 0.0000001D && Math.Abs(Right - right) < 0.0000001D &&
                bottom <= Top + tolerance && top >= Bottom - tolerance);

        internal void Join(double left, double bottom, double right, double top) {
            Left = Math.Min(Left, left); Bottom = Math.Min(Bottom, bottom);
            Right = Math.Max(Right, right); Top = Math.Max(Top, top);
        }
    }

    private static bool CanSelectGlyphs(PdfTextSpan span, out IReadOnlyList<PdfTextGlyph> glyphs) {
        glyphs = Array.Empty<PdfTextGlyph>();
        return span.ContentOrderKey is not null &&
            span.TextRenderingMode is >= 0 and <= 3 &&
            (!span.ClipPath.HasValue || span.ClipPath.Value.IsRectangle && span.ClipPath.Value.IsExact && !span.ClipPath.Value.ContainsTextClipping) && !span.IsType3Font &&
            string.Equals(span.Text, span.RestampText, StringComparison.Ordinal) && span.TryGetGlyphs(out glyphs) &&
            glyphs.All(static glyph => glyph.EncodedBytes.Count > 0 && glyph.PaintedAdvance > 0D &&
                !double.IsNaN(glyph.PaintedAdvance) && !double.IsInfinity(glyph.PaintedAdvance));
    }

    private static (int Page, PdfContentOrderKey Owner, int Run, int TextStart) GlyphKey(int page, PdfTextSpan span, PdfTextGlyph glyph) =>
        (page, span.ContentOrderKey!, span.GlyphSourceRunOrdinal, glyph.TextStart);
}

/// <summary>A complete-block selection before its glyphs are resolved: the mark rectangle and the spans that own it.</summary>
internal sealed class PdfRedactionBlockCandidate {
    internal PdfRedactionBlockCandidate(int pageNumber, PdfTextSpanBounds bounds, string label, IReadOnlyList<PdfTextSpan> owners, bool isHiddenOnly = false) {
        PageNumber = pageNumber; Bounds = bounds; Label = label; Owners = owners; IsHiddenOnly = isHiddenOnly;
    }

    internal int PageNumber { get; }
    internal PdfTextSpanBounds Bounds { get; }
    internal string Label { get; }
    internal IReadOnlyList<PdfTextSpan> Owners { get; }
    internal bool IsHiddenOnly { get; }
}

using System.Text.RegularExpressions;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfTextEditor {
    internal static bool HasRegexRedactionMatch(byte[] pdf, Regex expression, int pageNumber, PdfLoadOptions? readOptions,
        PdfRedactionSearchWorkBudget budget, CancellationToken cancellationToken) =>
        FindHits(pdf, expression.ToString(), new PdfTextSearchOptions { IncludeTextRenderingMode3 = true, PageNumbers = new[] { pageNumber } },
            readOptions, budget, unit => budget.Matches(expression, unit.Text)
                .Where(static match => match.Length > 0).Select(static match => new TextSearchRange(match.Index, match.Length, match.Value)),
            includeUnsafeSpans: true, cancellationToken: cancellationToken).Count > 0;

    // Reuse the editor's source-range, line-flow and table boundaries instead of
    // approximating a substring's position from logical-block text lengths.
    internal static IReadOnlyList<PdfRedactionArea> FindRedactionAreas(
        byte[] pdf, PdfRedactionSearchOptions search, Regex[] expressions, PdfLoadOptions? readOptions,
        PdfRedactionSearchWorkBudget budget, out IReadOnlyList<PdfDiagnosticFinding> findings) {
        var areas = new List<PdfRedactionArea>();
        var areaKeys = new HashSet<(int Page, double Left, double Bottom, double Right, double Top)>();
        var diagnostics = new List<PdfDiagnosticFinding>();
        var selected = new HashSet<(int Page, PdfContentOrderKey Owner, int Run, int TextStart)>();
        var diagnosticKeys = new HashSet<(int Page, string Code)>();
        var options = new PdfTextSearchOptions {
            MatchCase = search.MatchCase, IncludeTextRenderingMode3 = true,
            PageNumbers = search.PageNumbers.ToArray()
        };
        for (int index = 0; index < search.LiteralText.Count; index++) {
            string literal = search.LiteralText[index];
            AddHits(FindHits(pdf, literal, options, readOptions, budget, includeUnsafeSpans: true,
                cancellationToken: search.CancellationToken, maximumHits: search.MaximumCandidates), "literal:" + literal);
        }
        for (int index = 0; index < expressions.Length; index++) {
            Regex expression = expressions[index];
            AddHits(FindHits(pdf, expression.ToString(), options, readOptions, budget,
                unit => RegexRanges(unit.Text, expression), includeUnsafeSpans: true,
                cancellationToken: search.CancellationToken, maximumHits: search.MaximumCandidates),
                "regex:" + search.RegularExpressions[index]);
        }
        if (areas.Count > 0) {
            PdfReadDocument document = PdfReadDocument.Open(pdf, PdfLoadOptions.WithArtifactText(readOptions), search.CancellationToken);
            foreach (IGrouping<int, PdfRedactionArea> pageAreas in areas.GroupBy(static area => area.PageNumber)) {
                foreach (PdfTextSpan span in document.Pages[pageAreas.Key - 1].GetGlyphTextSpans(
                    includeHiddenOptionalContent: true, cancellationToken: search.CancellationToken)) {
                    search.CancellationToken.ThrowIfCancellationRequested();
                    if (!CanSelectGlyphs(span, out IReadOnlyList<PdfTextGlyph> glyphs)) {
                        PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetRedactionGlyphBounds(span, 0D, span.Advance);
                        if (IntersectsAreas(pageAreas, bounds)) Block(pageAreas.Key, "RedactionSearchGlyphMappingUnsupported",
                            "A precise search area intersects text without an unambiguous encoded glyph mapping.");
                        continue;
                    }
                    PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries);
                    foreach (PdfTextGlyph glyph in glyphs) {
                        budget.Charge(1L);
                        if (selected.Contains(GlyphKey(pageAreas.Key, span, glyph))) continue;
                        PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetRedactionGlyphBounds(span, boundaries[glyph.TextStart], glyph.PaintedAdvance);
                        if (IntersectsAreas(pageAreas, bounds)) Block(pageAreas.Key, "RedactionSearchUnselectedTextIntersection",
                            "A precise search area would remove an unselected glyph. Use reviewed geometry or complete-block selection instead.");
                    }
                }
            }
        }
        findings = diagnostics.AsReadOnly();
        return areas.AsReadOnly();

        IEnumerable<TextSearchRange> RegexRanges(string text, Regex expression) {
            foreach (Match match in budget.Matches(expression, text)) {
                search.CancellationToken.ThrowIfCancellationRequested();
                if (match.Length == 0) throw new InvalidOperationException("A precise redaction regex must select non-empty source text.");
                yield return new TextSearchRange(match.Index, match.Length, match.Value);
            }
        }

        void AddHits(IReadOnlyList<TextSearchHit> hits, string label) {
            foreach (TextSearchHit hit in hits) {
                foreach (TextSourceSegment segment in hit.Segments) {
                    search.CancellationToken.ThrowIfCancellationRequested();
                    PdfTextSpan span = segment.Span;
                    if (!CanSelectGlyphs(span, out IReadOnlyList<PdfTextGlyph> glyphs)) {
                        Block(hit.PageNumber, "RedactionSearchGlyphMappingUnsupported",
                            "The matched text does not have an unambiguous encoded glyph mapping. Precise selection cannot be reviewed.");
                        continue;
                    }
                    int end = segment.Start + segment.Length;
                    PdfTextGlyph[] covered = glyphs.Where(glyph => glyph.TextStart >= segment.Start && glyph.TextStart + glyph.TextLength <= end).ToArray();
                    if (covered.Length == 0 || covered[0].TextStart != segment.Start ||
                        covered[covered.Length - 1].TextStart + covered[covered.Length - 1].TextLength != end) {
                        Block(hit.PageNumber, "RedactionSearchPartialGlyph",
                            "The match selects only part of an encoded glyph, such as a ligature. Precise selection requires complete glyph boundaries.");
                        continue;
                    }
                    PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries);
                    double left = double.PositiveInfinity, bottom = double.PositiveInfinity;
                    double right = double.NegativeInfinity, top = double.NegativeInfinity;
                    foreach (PdfTextGlyph glyph in covered) {
                        budget.Charge(1L);
                        selected.Add(GlyphKey(hit.PageNumber, span, glyph));
                        PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetRedactionGlyphBounds(span, boundaries[glyph.TextStart], glyph.PaintedAdvance);
                        left = Math.Min(left, bounds.Left); bottom = Math.Min(bottom, bounds.Bottom);
                        right = Math.Max(right, bounds.Right); top = Math.Max(top, bounds.Top);
                    }
                    if (!areaKeys.Add((hit.PageNumber, left, bottom, right, top))) continue;
                    if (areas.Count >= search.MaximumCandidates) throw new InvalidOperationException("Redaction search exceeded the configured candidate limit.");
                    areas.Add(new PdfRedactionArea(hit.PageNumber, left, bottom, right - left, top - bottom, label, search.ContentScope)
                        .RequiringGlyphRewrite());
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

        void Block(int page, string code, string message) {
            if (diagnosticKeys.Add((page, code))) diagnostics.Add(new PdfDiagnosticFinding(PdfDiagnosticSeverity.Error, code, message, pageNumber: page));
        }
    }

    private static bool CanSelectGlyphs(PdfTextSpan span, out IReadOnlyList<PdfTextGlyph> glyphs) {
        glyphs = Array.Empty<PdfTextGlyph>();
        return span.ContentOrderKey is not null &&
            (!span.ClipPath.HasValue || span.ClipPath.Value.IsRectangle && span.ClipPath.Value.IsExact && !span.ClipPath.Value.ContainsTextClipping) && !span.IsType3Font &&
            string.Equals(span.Text, span.RestampText, StringComparison.Ordinal) && span.TryGetGlyphs(out glyphs) &&
            glyphs.All(static glyph => glyph.EncodedBytes.Count > 0 && glyph.PaintedAdvance > 0D &&
                !double.IsNaN(glyph.PaintedAdvance) && !double.IsInfinity(glyph.PaintedAdvance));
    }

    private static (int Page, PdfContentOrderKey Owner, int Run, int TextStart) GlyphKey(int page, PdfTextSpan span, PdfTextGlyph glyph) =>
        (page, span.ContentOrderKey!, span.GlyphSourceRunOrdinal, glyph.TextStart);
}

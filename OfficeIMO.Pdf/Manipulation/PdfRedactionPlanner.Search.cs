using System.Text.RegularExpressions;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionPlanner {
    /// <summary>Derives reviewable redaction rectangles from literal text, bounded regex, logical element kinds, and AcroForm field names.</summary>
    public static PdfRedactionPlan Search(byte[] pdf, PdfRedactionSearchOptions search, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf)); Guard.NotNull(search, nameof(search));
        search.CancellationToken.ThrowIfCancellationRequested();
        if (search.RegexTimeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(search), "Regex timeout must be positive.");
        if (search.MaximumCandidates <= 0) throw new ArgumentOutOfRangeException(nameof(search), "Maximum candidates must be positive.");
        if (search.LiteralText.Any(string.IsNullOrWhiteSpace))
            throw new ArgumentException("Literal search criteria must contain text.", nameof(search));
        Regex[] expressions = search.RegularExpressions.Select(pattern => new Regex(pattern, search.RegexOptions, search.RegexTimeout)).ToArray();
        if (search.LiteralText.Count == 0 && expressions.Length == 0 && search.FormFieldNames.Count == 0 && search.LogicalElementKinds.Count == 0) throw new ArgumentException("At least one redaction search criterion is required.", nameof(search));

        PdfReadDocument readDocument = PdfReadDocument.Open(pdf, readOptions, search.CancellationToken);
        PdfDocumentReadResult logical = PdfDocumentReadResult.From(readDocument, layoutOptions, search.CancellationToken);
        if (search.PageNumbers.Any(page => page < 1 || page > logical.Pages.Count)) throw new ArgumentOutOfRangeException(nameof(search), "Search pages must identify existing one-based pages.");
        search.CancellationToken.ThrowIfCancellationRequested();
        StringComparison comparison = search.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var areas = new List<PdfRedactionArea>(); var keys = new HashSet<string>(StringComparer.Ordinal);
        var workBudget = new PdfRedactionSearchWorkBudget("search planning");
        IReadOnlyList<PdfLogicalTextBlock> textBlocks = logical.TextBlocks;
        Dictionary<int, string> wrappedLiterals = MatchLiteralsAcrossBlocks(textBlocks, search.LiteralText, comparison, workBudget,
            index => search.PageNumbers.Count == 0 || search.PageNumbers.Contains(textBlocks[index].PageNumber), search.CancellationToken,
            readOptions?.Limits.MaxTextSearchFlowComparisons ?? PdfReadLimits.DefaultMaxTextSearchFlowComparisons);
        for (int blockIndex = 0; blockIndex < textBlocks.Count; blockIndex++) {
            PdfLogicalTextBlock block = textBlocks[blockIndex];
            search.CancellationToken.ThrowIfCancellationRequested();
            if (search.PageNumbers.Count > 0 && !search.PageNumbers.Contains(block.PageNumber)) continue;
            string? criterion = MatchText(block, search, expressions, comparison, workBudget) ??
                (wrappedLiterals.TryGetValue(blockIndex, out string? wrappedCriterion) ? wrappedCriterion : null);
            if (criterion is null) continue;
            PdfTextSpanBounds bounds = GetTextBlockBounds(block, logical.Pages[block.PageNumber - 1]);
            AddArea(areas, keys, new PdfRedactionArea(block.PageNumber, bounds.Left, bounds.Bottom, bounds.Width, bounds.Height, criterion), search.MaximumCandidates);
        }
        int[] tablePages = textBlocks.Where(static block => block.IsTableContent)
            .Select(static block => block.PageNumber)
            .Where(page => search.PageNumbers.Count == 0 || search.PageNumbers.Contains(page))
            .Distinct().ToArray();
        var embeddedBreakSpans = new Dictionary<int, PdfTextSpan[]>();
        if (search.LiteralText.Count > 0) {
            for (int pageNumber = 1; pageNumber <= readDocument.Pages.Count; pageNumber++) {
                if (search.PageNumbers.Count > 0 && !search.PageNumbers.Contains(pageNumber)) continue;
                PdfTextSpan[] withBreaks = readDocument.Pages[pageNumber - 1].GetTextSpans()
                    .Where(static span => span.EmbeddedLineBreaks?.Any(static isBreak => isBreak) == true).ToArray();
                if (withBreaks.Length > 0) embeddedBreakSpans[pageNumber] = withBreaks;
            }
        }
        int[] nativePages = tablePages.Concat(embeddedBreakSpans.Keys).Distinct().ToArray();
        if (nativePages.Length > 0 && search.LiteralText.Count > 0) {
            var nativeOptions = new PdfTextSearchOptions { MatchCase = search.MatchCase, IncludeTextRenderingMode3 = true, PageNumbers = nativePages };
            foreach (string literal in search.LiteralText) {
                search.CancellationToken.ThrowIfCancellationRequested();
                foreach (PdfTextMatch hit in PdfTextEditor.Find(pdf, literal, nativeOptions, readOptions, workBudget)) {
                    PdfReadPage page = readDocument.Pages[hit.PageNumber - 1];
                    PdfSelectionQuad visual = hit.VisualBounds;
                    PdfVisualBounds whole = page.TransformVisualBoundsToUser(visual.Left, visual.Top, visual.Right, visual.Bottom);
                    bool tableHit = hit.VisualLineBounds.Count > 1 &&
                        textBlocks.Any(block => block.PageNumber == hit.PageNumber && block.IsTableContent &&
                            Intersects(GetTextBlockBounds(block, logical.Pages[hit.PageNumber - 1]), whole));
                    bool embeddedBreakHit = embeddedBreakSpans.TryGetValue(hit.PageNumber, out PdfTextSpan[]? sourceSpans) &&
                        sourceSpans.Any(span => Intersects(PdfTextSpanGeometry.GetAxisAlignedBounds(span), whole));
                    if (!tableHit && !embeddedBreakHit) continue;
                    foreach (PdfSelectionQuad line in hit.VisualLineBounds) {
                        PdfVisualBounds bounds = page.TransformVisualBoundsToUser(line.Left, line.Top, line.Right, line.Bottom);
                        AddArea(areas, keys, new PdfRedactionArea(hit.PageNumber, bounds.Left, bounds.Top,
                            bounds.Width, bounds.Height, "literal:" + literal), search.MaximumCandidates);
                    }
                }
            }
        }
        var requestedFields = new HashSet<string>(search.FormFieldNames, StringComparer.Ordinal);
        foreach (PdfLogicalFormWidget widget in logical.FormWidgets) {
            search.CancellationToken.ThrowIfCancellationRequested();
            if (search.PageNumbers.Count > 0 && !search.PageNumbers.Contains(widget.PageNumber)) continue;
            if (widget.FieldName is not null && requestedFields.Contains(widget.FieldName)) AddArea(areas, keys, new PdfRedactionArea(widget.PageNumber, widget.X1, widget.Y1, widget.Width, widget.Height, "field:" + widget.FieldName), search.MaximumCandidates);
        }
        if (areas.Count == 0) return new PdfRedactionPlan(PdfInspector.Preflight(pdf, readOptions, search.CancellationToken), Array.Empty<PdfRedactionArea>(), Array.Empty<PdfRedactionMatch>(), new[] { new PdfDiagnosticFinding(PdfDiagnosticSeverity.Info, "RedactionSearchNoMatches", "No logical content matched the requested redaction search criteria.") }, DescribeCriteria(search), PdfRedactionPlan.ComputeSourceSha256(pdf), PdfRedactionPlan.CapturePageIdentities(readDocument, Array.Empty<PdfRedactionArea>()));
        PdfRedactionPlan planned = Plan(pdf, areas, layoutOptions, readOptions, search.CancellationToken);
        search.CancellationToken.ThrowIfCancellationRequested();
        return new PdfRedactionPlan(planned.Preflight, planned.Areas, planned.Matches, planned.Findings, DescribeCriteria(search), planned.SourceSha256, planned.PageIdentities, planned.ReviewedTextObjectScopes, search.MatchCase, search.RegexOptions, search.RegexTimeout);
    }

    private static string? MatchText(PdfLogicalTextBlock block, PdfRedactionSearchOptions search, Regex[] expressions, StringComparison comparison, PdfRedactionSearchWorkBudget workBudget) {
        for (int i = 0; i < search.LiteralText.Count; i++) {
            workBudget.ChargeTextScan(block.Text, search.LiteralText[i]);
            if (ContainsText(block.Text, search.LiteralText[i], comparison)) return "literal:" + search.LiteralText[i];
        }
        for (int i = 0; i < expressions.Length; i++) if (workBudget.IsMatch(expressions[i], block.Text)) return "regex:" + search.RegularExpressions[i];
        return search.LogicalElementKinds.Contains(block.Kind) ? "logical-kind:" + block.Kind.ToString() : null;
    }

    private static void AddArea(List<PdfRedactionArea> areas, HashSet<string> keys, PdfRedactionArea area, int maximumCandidates) { string key = area.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + area.X.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + ":" + area.Y.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + ":" + area.Width.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + ":" + area.Height.ToString("R", System.Globalization.CultureInfo.InvariantCulture); if (keys.Add(key)) { if (areas.Count >= maximumCandidates) throw new InvalidOperationException("Redaction search exceeded the configured candidate limit."); areas.Add(area); } }
    private static string[] DescribeCriteria(PdfRedactionSearchOptions search) => search.LiteralText.Select(value => "literal:" + value).Concat(search.RegularExpressions.Select(value => "regex:" + value)).Concat(search.FormFieldNames.Select(value => "field:" + value)).Concat(search.LogicalElementKinds.Select(value => "logical-kind:" + value.ToString())).ToArray();
    private static bool ContainsText(string text, string value, StringComparison comparison) =>
        PdfTextSearchNormalization.Contains(text, value, comparison);

    private static bool Intersects(PdfTextSpanBounds block, PdfVisualBounds match) =>
        block.Left < match.Right && block.Right > match.Left &&
        block.Bottom < match.Bottom && block.Top > match.Top;

    /// <summary>
    /// Finds literal occurrences across logical blocks in the same geometric text flow. The editor's
    /// line-flow builder owns rotation, column separation, and nearest-line ordering for both paths.
    /// Returns the criterion label for every participating block index.
    /// </summary>
    internal static Dictionary<int, string> MatchLiteralsAcrossBlocks(IReadOnlyList<PdfLogicalTextBlock> blocks, IEnumerable<string> literals,
        StringComparison comparison, PdfRedactionSearchWorkBudget workBudget, Func<int, bool> includeBlock, CancellationToken cancellationToken,
        int maxFlowComparisons = PdfReadLimits.DefaultMaxTextSearchFlowComparisons) {
        var matches = new Dictionary<int, string>();
        string[] requestedLiterals = literals as string[] ?? literals.ToArray();
        if (requestedLiterals.Length == 0) return matches;
        List<int[]> flows = BuildLogicalSearchFlows(blocks, includeBlock, maxFlowComparisons, cancellationToken);
        foreach (string literal in requestedLiterals) {
            int queryLength = PdfTextSearchNormalization.NormalizeQuery(literal).Length;
            foreach (int[] flow in flows) {
                for (int first = 0; first < flow.Length; first++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int firstIndex = flow[first];
                    if (string.IsNullOrEmpty(blocks[firstIndex].Text)) continue;
                    int firstEnd = blocks[firstIndex].Text.Length;
                    var combined = new System.Text.StringBuilder(blocks[firstIndex].Text);
                    int followingLength = 0;
                    for (int last = first + 1; last < flow.Length; last++) {
                        int lastIndex = flow[last];
                        combined.Append('\n');
                        int lastStart = combined.Length;
                        combined.Append(blocks[lastIndex].Text);
                        string joined = combined.ToString();
                        workBudget.ChargeTextScan(joined, literal);
                        bool spansRun = PdfTextSearchNormalization.FindSourceRanges(joined, literal, comparison)
                            .Any(range => range.Start < firstEnd && range.End > lastStart);
                        if (spansRun) {
                            for (int position = first; position <= last; position++) {
                                int index = flow[position];
                                if (!matches.ContainsKey(index)) matches.Add(index, "literal:" + literal);
                            }
                            break;
                        }
                        // A match starting in the first block and ending past this run contains every following
                        // block's visible characters (less at most one joined line-end hyphen).
                        followingLength += Math.Max(0, blocks[lastIndex].Text.Count(static character => !char.IsWhiteSpace(character)) - 1);
                        if (followingLength + 2 > queryLength) break;
                    }
                }
            }
        }
        return matches;
    }

    private static List<int[]> BuildLogicalSearchFlows(IReadOnlyList<PdfLogicalTextBlock> blocks,
        Func<int, bool> includeBlock, int maxFlowComparisons, CancellationToken cancellationToken) {
        var flows = new List<int[]>();
        var comparisonsByPage = new Dictionary<int, long>();
        foreach (IGrouping<(int PageNumber, PdfLogicalContentSourceKind SourceKind), int> group in Enumerable.Range(0, blocks.Count)
                     .Where(index => blocks[index].Spans.Count > 0 ||
                         blocks[index].FontSize > 0D && blocks[index].XEnd > blocks[index].XStart)
                     .GroupBy(index => (blocks[index].PageNumber, blocks[index].SourceKind))) {
            cancellationToken.ThrowIfCancellationRequested();
            int[] indexes = group.ToArray();
            if (!indexes.Any(includeBlock)) continue;
            var lines = indexes.Select(index => {
                PdfLogicalTextBlock block = blocks[index];
                List<PdfTextSpan> spans = block.Spans.Count > 0
                    ? block.Spans.ToList()
                    : new List<PdfTextSpan> { new(block.Text, string.Empty, block.FontSize,
                        block.XStart, block.BaselineY, block.XEnd - block.XStart) };
                return new TextLayoutEngine.TextLine(block.BaselineY, block.XStart, block.XEnd,
                    block.Text, spans);
            }).ToList();
            comparisonsByPage.TryGetValue(group.Key.PageNumber, out long comparisons);
            List<int[]> groupedFlows = PdfTextEditor.BuildSearchFlows(lines, maxFlowComparisons, ref comparisons);
            comparisonsByPage[group.Key.PageNumber] = comparisons;
            foreach (int[] flow in groupedFlows) {
                var eligible = new List<int>();
                foreach (int line in flow) {
                    int index = indexes[line];
                    if (!includeBlock(index) || blocks[index].IsTableContent) {
                        if (eligible.Count > 0) flows.Add(eligible.ToArray());
                        eligible.Clear();
                    } else {
                        eligible.Add(index);
                    }
                }
                if (eligible.Count > 0) flows.Add(eligible.ToArray());
            }
        }
        return flows;
    }
}

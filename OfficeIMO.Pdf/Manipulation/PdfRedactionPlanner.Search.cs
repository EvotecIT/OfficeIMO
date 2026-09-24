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
            index => search.PageNumbers.Count == 0 || search.PageNumbers.Contains(textBlocks[index].PageNumber), search.CancellationToken);
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

    /// <summary>
    /// Finds literal occurrences that wrap from one logical text block into the following blocks on the same page, such as
    /// a phrase broken across lines that layout analysis kept as separate blocks. Blocks are joined in reading order with a
    /// line break, and only runs whose match needs the first and the last block are reported. Returns the criterion label
    /// for every participating block index.
    /// </summary>
    internal static Dictionary<int, string> MatchLiteralsAcrossBlocks(IReadOnlyList<PdfLogicalTextBlock> blocks, IEnumerable<string> literals,
        StringComparison comparison, PdfRedactionSearchWorkBudget workBudget, Func<int, bool> includeBlock, CancellationToken cancellationToken) {
        var matches = new Dictionary<int, string>();
        foreach (string literal in literals) {
            int queryLength = PdfTextSearchNormalization.NormalizeQuery(literal).Length;
            for (int first = 0; first < blocks.Count; first++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!includeBlock(first) || string.IsNullOrEmpty(blocks[first].Text)) continue;
                int firstEnd = blocks[first].Text.Length;
                var combined = new System.Text.StringBuilder(blocks[first].Text);
                int followingLength = 0;
                for (int last = first + 1; last < blocks.Count && blocks[last].PageNumber == blocks[first].PageNumber && includeBlock(last); last++) {
                    combined.Append('\n');
                    int lastStart = combined.Length;
                    combined.Append(blocks[last].Text);
                    string joined = combined.ToString();
                    workBudget.ChargeTextScan(joined, literal);
                    bool spansRun = PdfTextSearchNormalization.FindSourceRanges(joined, literal, comparison)
                        .Any(range => range.Start < firstEnd && range.End > lastStart);
                    if (spansRun) {
                        for (int index = first; index <= last; index++) {
                            if (!matches.ContainsKey(index)) matches.Add(index, "literal:" + literal);
                        }
                        break;
                    }
                    // A match starting in the first block and ending past this run contains every following block's visible
                    // characters (less at most one joined line-end hyphen), plus at least one character on each side.
                    followingLength += Math.Max(0, blocks[last].Text.Count(static character => !char.IsWhiteSpace(character)) - 1);
                    if (followingLength + 2 > queryLength) break;
                }
            }
        }
        return matches;
    }
}

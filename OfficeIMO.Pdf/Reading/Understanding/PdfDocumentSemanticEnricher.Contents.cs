using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfDocumentSemanticEnricher {
    private readonly record struct TaggedContentsContent(int ObjectNumber, int Level);

    /// <summary>
    /// Recognises contents entries, from tagged TOC/TOCI structure or from the untagged contents-line pattern
    /// (title, a dot leader or a wide gap, and a trailing page number on consecutive lines). Table candidates
    /// that would split the entries into columns are withdrawn so each entry is projected once, as a list item.
    /// </summary>
    private static IReadOnlyList<PdfContentsEntrySource>[] BuildContentsSources(
        PdfReadDocument document,
        int[] pageNumbers,
        IReadOnlyList<PdfUnderstandingPageResult> pages,
        List<PdfUnderstandingSemanticElement>[] elements,
        IReadOnlyList<PdfUnderstandingTableCandidate>[] tableCandidates,
        TaggedContentRoleIndex? taggedRoles,
        int maximum,
        PdfUnderstandingWorkBudget budget) {
        var result = new IReadOnlyList<PdfContentsEntrySource>[pages.Count];
        HashSet<string>? knownTitles = null;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            budget.Consume();
            int pageNumber = pageNumbers[pageIndex];
            PdfReadPage readPage = document.Pages[pageNumber - 1];
            PdfUnderstandingPageResult page = pages[pageIndex];
            List<PdfUnderstandingLine> lines = page.Lines
                .Where(static line => line.Words.Count > 0 && Math.Abs(line.RotationDegrees) < 0.5D)
                .OrderByDescending(static line => line.BaselineY)
                .ThenBy(static line => line.XStart)
                .ToList();
            var entries = new List<PdfContentsEntrySource>();
            var claimed = new HashSet<PdfUnderstandingLine>();
            if (taggedRoles is not null) AddTaggedContentsEntries(pageNumber, readPage, lines, taggedRoles, entries, claimed, budget);
            knownTitles ??= CollectKnownTitles(document, elements, budget);
            AddPatternContentsEntries(document, readPage, pageNumber, lines, taggedRoles, knownTitles, entries, claimed, budget);
            if (entries.Count > maximum) throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maximum, entries.Count);
            if (entries.Count > 0) {
                var owned = new HashSet<PdfTextSpan>(entries.SelectMany(static entry => entry.Runs));
                IReadOnlyList<PdfUnderstandingTableCandidate> candidates = tableCandidates[pageIndex];
                PdfUnderstandingTableCandidate[] retained = candidates
                    .Where(candidate => !candidate.SourceLines.Any(line => line.Words.Any(word => word.SourceRuns.Any(owned.Contains))))
                    .ToArray();
                if (retained.Length != candidates.Count) tableCandidates[pageIndex] = retained;
            }
            result[pageIndex] = entries.Count == 0 ? Array.Empty<PdfContentsEntrySource>() : entries.AsReadOnly();
        }
        return result;
    }

    private static void AddTaggedContentsEntries(
        int pageNumber,
        PdfReadPage readPage,
        List<PdfUnderstandingLine> lines,
        TaggedContentRoleIndex roles,
        List<PdfContentsEntrySource> entries,
        HashSet<PdfUnderstandingLine> claimed,
        PdfUnderstandingWorkBudget budget) {
        var owners = new Dictionary<int, (int Level, List<PdfUnderstandingWord> Words, List<PdfUnderstandingLine> Lines)>();
        var order = new List<int>();
        foreach (PdfUnderstandingLine line in lines) {
            foreach (PdfUnderstandingWord word in line.Words) {
                budget.Consume();
                TaggedContentsContent? owner = null;
                foreach (PdfTextSpan run in word.SourceRuns) {
                    if (!run.MarkedContentId.HasValue) continue;
                    owner ??= roles.GetContentsContent(pageNumber, readPage,
                        new MarkedContentKey(run.ContentStreamObjectNumber, run.MarkedContentId.Value));
                }
                if (!owner.HasValue) continue;
                if (!owners.TryGetValue(owner.Value.ObjectNumber, out var item)) {
                    item = (owner.Value.Level, new List<PdfUnderstandingWord>(), new List<PdfUnderstandingLine>());
                    owners.Add(owner.Value.ObjectNumber, item);
                    order.Add(owner.Value.ObjectNumber);
                }
                item.Words.Add(word);
                if (!item.Lines.Contains(line)) item.Lines.Add(line);
            }
        }

        foreach (int objectNumber in order) {
            budget.Consume();
            var item = owners[objectNumber];
            if (!TrySplitContentsEntry(item.Words, strictPageLabel: false, out List<PdfUnderstandingWord> title, out string? pageLabel, out _)) continue;
            entries.Add(new PdfContentsEntrySource(item.Level, title, pageLabel, item.Words, item.Lines,
                new PdfInferenceEvidence("semantic-tagged-contents-item",
                    "The nearest tagged TOCI structure " + objectNumber.ToString(CultureInfo.InvariantCulture) +
                    " owns this contents entry's text on the page.", 0.99D)));
            foreach (PdfUnderstandingLine line in item.Lines) claimed.Add(line);
        }
    }

    private static void AddPatternContentsEntries(
        PdfReadDocument document,
        PdfReadPage readPage,
        int pageNumber,
        List<PdfUnderstandingLine> lines,
        TaggedContentRoleIndex? roles,
        HashSet<string> knownTitles,
        List<PdfContentsEntrySource> entries,
        HashSet<PdfUnderstandingLine> claimed,
        PdfUnderstandingWorkBudget budget) {
        IReadOnlyList<PdfLinkAnnotation>? links = null;
        List<ContentsRow> rows = GroupContentsRows(lines, budget);
        var run = new List<PatternContentsLine>();
        for (int rowIndex = 0; rowIndex <= rows.Count; rowIndex++) {
            budget.Consume();
            PatternContentsLine? candidate = null;
            if (rowIndex < rows.Count && !rows[rowIndex].Lines.Any(claimed.Contains) &&
                !rows[rowIndex].Lines.Any(line => HasTaggedListOrTableRole(line, pageNumber, readPage, roles)) &&
                TrySplitContentsEntry(rows[rowIndex].Words, strictPageLabel: true,
                    out List<PdfUnderstandingWord> title, out string? pageLabel, out bool hasLeader) &&
                pageLabel is not null) {
                candidate = new PatternContentsLine(rows[rowIndex], title, pageLabel, hasLeader);
            }
            // A contents entry's page number is right-aligned with the entry above it.
            if (candidate is not null && (run.Count == 0 ||
                Math.Abs(run[run.Count - 1].Row.XEnd - candidate.Row.XEnd) <= Math.Max(3D, candidate.Row.FontSize * 0.5D))) {
                run.Add(candidate);
                continue;
            }
            AcceptPatternRun();
            run.Clear();
            if (candidate is not null) run.Add(candidate);
        }

        void AcceptPatternRun() {
            if (run.Count < 2 || !HasPlausiblePageSequence(run, document.Pages.Count)) return;
            links ??= readPage.GetLinkAnnotations();
            int corroborated = run.Count(line => HasInternalLink(line, links) || knownTitles.Contains(NormalizeTitle(line.Title)));
            // Leader rows such as price lists share the shape; contents entries also point into the document.
            if (corroborated * 2 < run.Count) return;
            double[] indents = run.Select(static line => line.Title[0].XStart).Distinct().OrderBy(static x => x).ToArray();
            var levels = new List<double>();
            foreach (double indent in indents) {
                if (levels.Count == 0 || indent - levels[levels.Count - 1] > 3D) levels.Add(indent);
            }
            foreach (PatternContentsLine line in run) {
                budget.Consume();
                int level = levels.FindLastIndex(indent => line.Title[0].XStart >= indent - 3D) + 1;
                entries.Add(new PdfContentsEntrySource(Math.Max(1, level), line.Title, line.PageLabel, line.Row.Words,
                    line.Row.Lines,
                    new PdfInferenceEvidence("contents-line-pattern",
                        "Consecutive lines end in right-aligned page numbers after " + (line.HasLeader ? "a dot leader" : "a wide gap") +
                        " and point into the document.", 0.85D)));
                foreach (PdfUnderstandingLine source in line.Row.Lines) claimed.Add(source);
            }
        }
    }

    /// <summary>Joins line fragments that share a baseline, since a wide gap before a page number can split one row.</summary>
    private static List<ContentsRow> GroupContentsRows(List<PdfUnderstandingLine> lines, PdfUnderstandingWorkBudget budget) {
        var rows = new List<ContentsRow>();
        foreach (PdfUnderstandingLine line in lines) {
            budget.Consume();
            ContentsRow? row = rows.Count == 0 ? null : rows[rows.Count - 1];
            if (row is not null && Math.Abs(row.BaselineY - line.BaselineY) <= Math.Max(1D, Math.Min(row.FontSize, line.FontSize) * 0.3D)) {
                row.Add(line);
            } else {
                rows.Add(new ContentsRow(line));
            }
        }
        return rows;
    }

    private static bool HasTaggedListOrTableRole(PdfUnderstandingLine line, int pageNumber, PdfReadPage readPage, TaggedContentRoleIndex? roles) {
        if (roles is null || !roles.HasPage(pageNumber)) return false;
        foreach (PdfUnderstandingWord word in line.Words) {
            foreach (PdfTextSpan run in word.SourceRuns) {
                if (!run.MarkedContentId.HasValue) continue;
                IReadOnlyList<string> found = roles.GetRoles(pageNumber, readPage, new MarkedContentKey(run.ContentStreamObjectNumber, run.MarkedContentId.Value));
                if (found.Any(static role => role is "Table" or "TR" or "TD" or "TH" or "L" or "LI")) return true;
            }
        }
        return false;
    }

    private static bool HasPlausiblePageSequence(List<PatternContentsLine> run, int pageCount) {
        bool roman = TryParseRoman(run[0].PageLabel, out _);
        int previous = 0;
        foreach (PatternContentsLine line in run) {
            int value;
            if (roman) {
                if (!TryParseRoman(line.PageLabel, out value)) return false;
            } else if (!int.TryParse(line.PageLabel, NumberStyles.None, CultureInfo.InvariantCulture, out value) || value < 1 || value > pageCount) {
                return false;
            }
            if (value < previous) return false;
            previous = value;
        }
        return true;
    }

    private static bool HasInternalLink(PatternContentsLine line, IReadOnlyList<PdfLinkAnnotation> links) {
        double top = line.Row.BaselineY + line.Row.FontSize * 0.8D;
        double bottom = line.Row.BaselineY - line.Row.FontSize * 0.3D;
        foreach (PdfLinkAnnotation link in links) {
            if (link.Uri is not null || (link.DestinationName is null && !link.DestinationPageNumber.HasValue)) continue;
            if (Math.Min(link.Y1, link.Y2) <= top && Math.Max(link.Y1, link.Y2) >= bottom &&
                Math.Min(link.X1, link.X2) <= line.Row.XEnd && Math.Max(link.X1, link.X2) >= line.Row.XStart) return true;
        }
        return false;
    }

    private static HashSet<string> CollectKnownTitles(
        PdfReadDocument document,
        List<PdfUnderstandingSemanticElement>[] elements,
        PdfUnderstandingWorkBudget budget) {
        var titles = new HashSet<string>(StringComparer.Ordinal);
        foreach (PdfOutlineItem outline in FlattenOutlines(document.Outlines)) {
            budget.Consume();
            if (!string.IsNullOrWhiteSpace(outline.Title)) titles.Add(NormalizeTitle(outline.Title));
        }
        foreach (List<PdfUnderstandingSemanticElement> page in elements) {
            foreach (PdfUnderstandingSemanticElement element in page) {
                budget.Consume();
                if (element.Kind == PdfUnderstandingSemanticKind.Heading) titles.Add(NormalizeTitle(element.Region.Text));
            }
        }
        return titles;
    }

    private static string NormalizeTitle(IEnumerable<PdfUnderstandingWord> words) => NormalizeTitle(string.Join(" ", words.Select(static word => word.Text)));

    private static string NormalizeTitle(string text) =>
        string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).ToUpperInvariant();

    /// <summary>
    /// Splits an entry into title words and a trailing page label. The label must follow a dot leader or a gap
    /// of at least one and a half text heights, so a title that merely ends in a number keeps it.
    /// </summary>
    private static bool TrySplitContentsEntry(
        List<PdfUnderstandingWord> words,
        bool strictPageLabel,
        out List<PdfUnderstandingWord> title,
        out string? pageLabel,
        out bool hasLeader) {
        title = new List<PdfUnderstandingWord>();
        pageLabel = null;
        hasLeader = false;
        int end = words.Count;
        if (end == 0) return false;
        PdfUnderstandingWord last = words[end - 1];
        string labelText = last.Text.Trim();
        int labelLeader = CountLeading(labelText);
        string candidateLabel = labelText.Substring(labelLeader);
        if (IsPageLabel(candidateLabel, strictPageLabel) && end > 1) {
            int titleEnd = end - 1;
            bool leader = labelLeader >= 3;
            while (titleEnd > 0 && IsLeaderWord(words[titleEnd - 1].Text)) {
                leader = true;
                titleEnd--;
            }
            if (titleEnd > 0) {
                PdfUnderstandingWord previous = words[titleEnd - 1];
                bool trailingLeader = CountTrailing(previous.Text.Trim()) >= 3;
                double gap = last.XStart - words[end - 2].XEnd;
                if (leader || trailingLeader || gap >= Math.Max(last.FontSize, previous.FontSize) * 1.5D) {
                    hasLeader = leader || trailingLeader;
                    pageLabel = candidateLabel;
                    end = titleEnd;
                }
            }
        }
        if (strictPageLabel && pageLabel is null) return false;
        for (int index = 0; index < end; index++) {
            if (IsLeaderWord(words[index].Text)) continue;
            title.Add(words[index]);
        }
        return title.Count > 0 && title.Any(static word => word.Text.Any(char.IsLetter));
    }

    internal static string TrimContentsLeader(string text) {
        string trimmed = text.Trim();
        int trailing = CountTrailing(trimmed);
        return trailing >= 3 ? trimmed.Substring(0, trimmed.Length - trailing).TrimEnd() : trimmed;
    }

    // Dot, hyphen and underscore leaders, plus the middle-dot and ellipsis forms other producers paint.
    private static bool IsLeaderCharacter(char ch) => ch is '.' or '-' or '_' or '\u00B7' or '\u2026' or '\u2022' or '\u2219' or '\u22C5';

    private static bool IsLeaderWord(string text) {
        string trimmed = text.Trim();
        return trimmed.Length > 0 && (trimmed.Length >= 3 || trimmed[0] == '\u2026') && trimmed.All(IsLeaderCharacter);
    }

    private static int CountLeading(string text) {
        int count = 0;
        while (count < text.Length && IsLeaderCharacter(text[count])) count++;
        return count;
    }

    private static int CountTrailing(string text) {
        int count = 0;
        while (count < text.Length && IsLeaderCharacter(text[text.Length - 1 - count])) count++;
        return count;
    }

    private static bool IsPageLabel(string text, bool strict) {
        if (text.Length == 0 || text.Length > 8) return false;
        if (text.Length <= 5 && text.All(static ch => ch >= '0' && ch <= '9')) return true;
        if (TryParseRoman(text, out _)) return true;
        // Formatted labels such as "p3" or "A-2" are accepted only when tagged structure owns the entry.
        return !strict && text.Any(char.IsDigit) && !text.Any(char.IsWhiteSpace);
    }

    private static bool TryParseRoman(string text, out int value) {
        value = 0;
        if (text.Length == 0 || text.Length > 12) return false;
        bool upper = char.IsUpper(text[0]);
        for (int index = 0; index < text.Length; index++) {
            char ch = text[index];
            if (char.IsUpper(ch) != upper) return false;
            int digit = char.ToUpperInvariant(ch) switch {
                'I' => 1, 'V' => 5, 'X' => 10, 'L' => 50, 'C' => 100, 'D' => 500, 'M' => 1000, _ => 0
            };
            if (digit == 0) return false;
            int next = index + 1 < text.Length ? char.ToUpperInvariant(text[index + 1]) switch {
                'I' => 1, 'V' => 5, 'X' => 10, 'L' => 50, 'C' => 100, 'D' => 500, 'M' => 1000, _ => 0
            } : 0;
            if (next > digit) {
                value -= digit;
            } else {
                value += digit;
            }
        }
        return value > 0 && string.Equals(ToRoman(value), text.ToUpperInvariant(), StringComparison.Ordinal);
    }

    private static string ToRoman(int value) {
        int[] numbers = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        string[] symbols = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        var builder = new StringBuilder();
        for (int index = 0; index < numbers.Length; index++) {
            while (value >= numbers[index]) {
                builder.Append(symbols[index]);
                value -= numbers[index];
            }
        }
        return builder.ToString();
    }

    private sealed class ContentsRow {
        internal ContentsRow(PdfUnderstandingLine line) {
            BaselineY = line.BaselineY;
            FontSize = line.FontSize;
            Add(line);
        }

        internal double BaselineY { get; }
        internal double FontSize { get; }
        internal List<PdfUnderstandingLine> Lines { get; } = new();
        internal List<PdfUnderstandingWord> Words { get; private set; } = new();
        internal double XStart => Words[0].XStart;
        internal double XEnd => Words[Words.Count - 1].XEnd;

        internal void Add(PdfUnderstandingLine line) {
            Lines.Add(line);
            Words = Words.Concat(line.Words).OrderBy(static word => word.XStart).ToList();
        }
    }

    private sealed class PatternContentsLine {
        internal PatternContentsLine(ContentsRow row, List<PdfUnderstandingWord> title, string pageLabel, bool hasLeader) {
            Row = row;
            Title = title;
            PageLabel = pageLabel;
            HasLeader = hasLeader;
        }

        internal ContentsRow Row { get; }
        internal List<PdfUnderstandingWord> Title { get; }
        internal string PageLabel { get; }
        internal bool HasLeader { get; }
    }

    private sealed partial class TaggedContentRoleIndex {
        private readonly Dictionary<int, Dictionary<MarkedContentKey, TaggedContentsContent?>> _contentsByPage = new();

        private void AddContentsContent(int pageNumber, MarkedContentKey key, TaggedContentsContent? content) {
            if (!content.HasValue) return;
            if (!_contentsByPage.TryGetValue(pageNumber, out Dictionary<MarkedContentKey, TaggedContentsContent?>? page)) {
                page = new Dictionary<MarkedContentKey, TaggedContentsContent?>();
                _contentsByPage.Add(pageNumber, page);
            }
            // Conflicting structure references cannot establish a unique contents entry.
            if (page.TryGetValue(key, out TaggedContentsContent? existing)) {
                if (existing != content) page[key] = null;
            } else {
                page.Add(key, content);
            }
        }

        internal TaggedContentsContent? GetContentsContent(int pageNumber, PdfReadPage readPage, MarkedContentKey key) {
            if (!_contentsByPage.TryGetValue(pageNumber, out Dictionary<MarkedContentKey, TaggedContentsContent?>? page)) return null;
            if (page.TryGetValue(key, out TaggedContentsContent? exact)) return exact;
            if (!key.ContentStreamObjectNumber.HasValue) {
                MarkedContentKey? scoped = ResolveUniquePageContentKey(pageNumber, readPage, key.MarkedContentId);
                return scoped.HasValue && page.TryGetValue(scoped.Value, out TaggedContentsContent? found) ? found : null;
            }
            return readPage.IsPageContentStreamObjectNumber(key.ContentStreamObjectNumber) &&
                page.TryGetValue(new MarkedContentKey(null, key.MarkedContentId), out TaggedContentsContent? unscoped)
                ? unscoped : null;
        }
    }
}

/// <summary>One contents entry: its level, title words, printed page label and the source text it owns.</summary>
internal sealed class PdfContentsEntrySource {
    internal PdfContentsEntrySource(
        int level,
        IReadOnlyList<PdfUnderstandingWord> title,
        string? pageLabel,
        IReadOnlyList<PdfUnderstandingWord> words,
        IReadOnlyList<PdfUnderstandingLine> lines,
        PdfInferenceEvidence evidence) {
        Level = level;
        Title = title;
        PageLabel = pageLabel;
        Lines = lines;
        Evidence = evidence;
        Runs = new HashSet<PdfTextSpan>(words.SelectMany(static word => word.SourceRuns));
    }

    internal int Level { get; }
    internal IReadOnlyList<PdfUnderstandingWord> Title { get; }
    internal string? PageLabel { get; }
    internal IReadOnlyList<PdfUnderstandingLine> Lines { get; }
    internal HashSet<PdfTextSpan> Runs { get; }
    internal PdfInferenceEvidence Evidence { get; }
}

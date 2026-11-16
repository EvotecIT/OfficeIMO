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
        ContentsContinuation? continuation = null;
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
            if (pageIndex > 0 && pageNumbers[pageIndex - 1] != pageNumber - 1) continuation = null;
            AddPatternContentsEntries(document, readPage, pageNumber, lines, taggedRoles, knownTitles, entries, claimed, ref continuation, budget);
            if (entries.Count > maximum) throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maximum, entries.Count);
            // Columns painted by one text operation share a source run, so neither entry owns its text alone.
            // Leave such lines as they were read rather than assign the run to one of them.
            var runOwners = new Dictionary<PdfTextSpan, int>();
            foreach (PdfTextSpan run in entries.SelectMany(static entry => entry.Runs)) {
                runOwners[run] = runOwners.TryGetValue(run, out int count) ? count + 1 : 1;
            }
            entries.RemoveAll(entry => entry.Runs.Any(run => runOwners[run] > 1));
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
            var ownedWords = new HashSet<PdfUnderstandingWord>(item.Words);
            item.Words.Clear();
            foreach (ContentsRow row in GroupContentsRows(item.Lines, budget)) item.Words.AddRange(row.Words.Where(ownedWords.Contains));
            List<PdfUnderstandingWord> title;
            string? pageLabel;
            string? lastTitleText = null;
            // Structure decides first: the title is the entry's Link and the page number is the text after it.
            // With an artifact leader the gap can be narrow, and the title can end in its own dots.
            int linkEnd = item.Words.FindLastIndex(word => HasRole(word, "Link"));
            List<PdfUnderstandingWord> afterLink = item.Words.Skip(linkEnd + 1).Where(word => !IsLeaderWord(word.Text)).ToList();
            string structuralLabel = string.Join(" ", afterLink.Select(static word => word.Text.Trim()));
            // A formatted label such as "Page 2" is a few words; anything longer is not a page reference.
            if (linkEnd >= 0 && item.Words.Take(linkEnd + 1).All(word => HasRole(word, "Link")) &&
                afterLink.Count is > 0 and <= 3 && structuralLabel.Length <= 24 &&
                IsPageLabel(afterLink[afterLink.Count - 1].Text.Trim(), strict: false)) {
                title = item.Words.Take(linkEnd + 1).ToList();
                // A leader painted inside the link is still a separator, not title text. An ellipsis such as
                // "..." or "\u2026" is shorter than any leader and stays with the title.
                while (title.Count > 0 && IsLeaderWord(title[title.Count - 1].Text) && title[title.Count - 1].Text.Trim().Length > 3) {
                    title.RemoveAt(title.Count - 1);
                }
                pageLabel = structuralLabel;
            } else if (!TrySplitContentsEntry(item.Words, strictPageLabel: false, out title, out pageLabel, out _, out lastTitleText)) {
                continue;
            }
            if (!title.Any(static word => word.Text.Any(char.IsLetter))) continue;
            entries.Add(new PdfContentsEntrySource(item.Level, title, pageLabel, item.Words, item.Lines,
                new PdfInferenceEvidence("semantic-tagged-contents-item",
                    "The nearest tagged TOCI structure " + objectNumber.ToString(CultureInfo.InvariantCulture) +
                    " owns this contents entry's text on the page.", 0.99D), lastTitleText));
            foreach (PdfUnderstandingLine line in item.Lines) claimed.Add(line);
        }

        bool HasRole(PdfUnderstandingWord word, string role) => word.SourceRuns.Any(run => run.MarkedContentId.HasValue &&
            roles.GetRoles(pageNumber, readPage, new MarkedContentKey(run.ContentStreamObjectNumber, run.MarkedContentId.Value))
                .Contains(role, StringComparer.OrdinalIgnoreCase));
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
        ref ContentsContinuation? continuation,
        PdfUnderstandingWorkBudget budget) {
        IReadOnlyList<PdfLinkAnnotation>? links = null;
        List<ContentsRow> rows = GroupContentsRows(lines, budget);
        var open = new List<List<PatternContentsLine>>();
        var completed = new List<List<PatternContentsLine>>();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            budget.Consume();
            ContentsRow row = rows[rowIndex];
            if (row.Lines.Any(claimed.Contains) || row.Lines.Any(line => HasTaggedListOrTableRole(line, pageNumber, readPage, roles))) {
                CloseAll();
                continue;
            }
            var next = new List<List<PatternContentsLine>>();
            foreach (List<PdfUnderstandingWord> segment in SplitContentsColumns(row.Words)) {
                if (!TrySplitContentsEntry(segment, strictPageLabel: true,
                        out List<PdfUnderstandingWord> title, out string? pageLabel, out bool hasLeader, out string? lastTitleText) || pageLabel is null) continue;
                var candidate = new PatternContentsLine(row, rowIndex, segment, title, pageLabel, hasLeader, lastTitleText);
                // A contents entry's page number is right-aligned with the entry above it in the same column.
                List<PatternContentsLine>? run = open.FirstOrDefault(item =>
                    Math.Abs(item[item.Count - 1].XEnd - candidate.XEnd) <= Math.Max(3D, row.FontSize * 0.5D));
                if (run is null) {
                    run = new List<PatternContentsLine>();
                } else {
                    open.Remove(run);
                }
                run.Add(candidate);
                next.Add(run);
            }
            // Columns can sit on offset baselines, so a run only ends when other text crosses its column.
            foreach (List<PatternContentsLine> run in open) {
                double left = run.Min(static line => line.XStart);
                double right = run.Max(static line => line.XEnd);
                if (row.Words.Any(word => word.XStart < right && word.XEnd > left)) {
                    completed.Add(run);
                } else {
                    next.Add(run);
                }
            }
            open = next;
        }
        CloseAll();

        ContentsContinuation? previousPage = continuation;
        var acceptedRuns = new List<List<PatternContentsLine>>();
        bool continuesPreviousPage = false;
        // One entry alone matches too much. It counts only beside accepted contents on the page (another column)
        // or as the tail of contents from the previous page.
        foreach (bool singles in new[] { false, true }) {
            foreach (List<PatternContentsLine> run in completed.OrderBy(static item => item[0].RowIndex)) {
                if ((run.Count < 2) != singles) continue;
                Accept(run);
            }
        }
        if (acceptedRuns.Count == 0) {
            continuation = null;
            return;
        }

        // Accepted columns, and contents continued from the previous page, share one indentation scale.
        double[] indents = acceptedRuns.SelectMany(static run => run.Select(static line => line.Indent))
            .Concat(continuesPreviousPage ? previousPage!.Indents : Array.Empty<double>())
            .Distinct().OrderByDescending(static indent => indent).ToArray();
        var levels = new List<double>();
        foreach (double indent in indents) {
            if (levels.Count == 0 || levels[levels.Count - 1] - indent > 3D) levels.Add(indent);
        }
        PageLabelValue? lastValue = null;
        foreach (List<PatternContentsLine> run in acceptedRuns) {
            foreach (PatternContentsLine line in run) {
                budget.Consume();
                int level = levels.FindLastIndex(indent => line.Indent <= indent + 3D) + 1;
                PdfUnderstandingLine[] sources = line.Row.Lines.Where(source => source.Words.Any(line.Words.Contains)).ToArray();
                entries.Add(new PdfContentsEntrySource(Math.Max(1, level), line.Title, line.PageLabel, line.Words, sources,
                    new PdfInferenceEvidence("contents-line-pattern",
                        "Consecutive lines end in right-aligned page numbers after " + (line.HasLeader ? "a dot leader" : "a wide gap") +
                        " and point into the document.", 0.85D), line.LastTitleText));
                foreach (PdfUnderstandingLine source in sources) claimed.Add(source);
            }
            PageLabelValue? end = TryReadPageValue(run[run.Count - 1].PageLabel, document.Pages.Count);
            if (end.HasValue && (lastValue is not PageLabelValue known || known.IsBefore(end.Value))) {
                lastValue = end;
            }
        }
        continuation = lastValue.HasValue ? new ContentsContinuation(lastValue, levels.ToArray()) : null;

        void Accept(List<PatternContentsLine> run) {
            budget.Consume();
            PageLabelValue? first = TryReadPageSequence(run, document.Pages.Count);
            if (!first.HasValue) return;
            // The column that continues the previous page's contents keeps its indentation levels.
            // Indents are measured from the page numbers, so this holds when the next page starts in another column.
            bool continues = previousPage?.Last is PageLabelValue previousLast && !first.Value.IsBefore(previousLast);
            if (run.Count < 2 && acceptedRuns.Count == 0 && !continues) return;
            links ??= readPage.GetLinkAnnotations();
            int corroborated = run.Count(line => HasInternalLink(line, links) || knownTitles.Contains(NormalizeTitle(line.DisplayTitle)));
            // Leader rows such as price lists share the shape; contents entries also point into the document.
            if (corroborated * 2 < run.Count) return;
            acceptedRuns.Add(run);
            continuesPreviousPage |= continues;
        }

        void CloseAll() {
            completed.AddRange(open);
            open = new List<List<PatternContentsLine>>();
        }
    }

    /// <summary>Splits a baseline row after each page number that is followed by a wide gap, so side-by-side contents columns stay apart.</summary>
    private static IEnumerable<List<PdfUnderstandingWord>> SplitContentsColumns(List<PdfUnderstandingWord> words) {
        var segment = new List<PdfUnderstandingWord>();
        for (int index = 0; index < words.Count; index++) {
            PdfUnderstandingWord word = words[index];
            segment.Add(word);
            // "Chapter 1    2" must not end a column after "1": the segment has to be a whole entry already.
            bool endsColumn = index + 1 < words.Count &&
                words[index + 1].XStart - word.XEnd >= word.FontSize * 1.5D &&
                TrySplitContentsEntry(segment, strictPageLabel: true, out _, out string? label, out _, out _) && label is not null;
            if (endsColumn) {
                yield return segment;
                segment = new List<PdfUnderstandingWord>();
            }
        }
        if (segment.Count > 0) yield return segment;
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
        foreach (ContentsRow row in rows) row.Complete();
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

    /// <summary>
    /// Reads the run's page numbers: front-matter roman numerals may be followed by arabic numbers, each section
    /// in order, and arabic numbers within the page count. Returns the first value, or null when the run does not fit.
    /// </summary>
    private static PageLabelValue? TryReadPageSequence(List<PatternContentsLine> run, int pageCount) {
        PageLabelValue? first = null;
        PageLabelValue? previous = null;
        foreach (PatternContentsLine line in run) {
            PageLabelValue? value = TryReadPageValue(line.PageLabel, pageCount);
            if (!value.HasValue || (previous.HasValue && value.Value.IsBefore(previous.Value))) return null;
            first ??= value;
            previous = value;
        }
        return first;
    }

    private static PageLabelValue? TryReadPageValue(string label, int pageCount) {
        if (TryParseRoman(label, out int roman)) return new PageLabelValue(true, roman);
        return int.TryParse(label, NumberStyles.None, CultureInfo.InvariantCulture, out int value) && value >= 1 && value <= pageCount
            ? new PageLabelValue(false, value)
            : null;
    }

    /// <summary>Where the previous page's contents ended: its last page number and indentation levels.</summary>
    private sealed class ContentsContinuation {
        internal ContentsContinuation(PageLabelValue? last, double[] indents) {
            Last = last;
            Indents = indents;
        }

        internal PageLabelValue? Last { get; }
        internal double[] Indents { get; }
    }

    private readonly record struct PageLabelValue(bool Roman, int Value) {
        // Roman front matter precedes the arabic body numbering.
        internal bool IsBefore(PageLabelValue other) => Roman != other.Roman ? Roman : Value < other.Value;
    }

    private static bool HasInternalLink(PatternContentsLine line, IReadOnlyList<PdfLinkAnnotation> links) {
        double top = line.Row.BaselineY + line.Row.FontSize * 0.8D;
        double bottom = line.Row.BaselineY - line.Row.FontSize * 0.3D;
        foreach (PdfLinkAnnotation link in links) {
            if (link.Uri is not null || (link.DestinationName is null && !link.DestinationPageNumber.HasValue && !link.DestinationPageObjectNumber.HasValue)) continue;
            if (Math.Min(link.Y1, link.Y2) <= top && Math.Max(link.Y1, link.Y2) >= bottom &&
                Math.Min(link.X1, link.X2) <= line.XEnd && Math.Max(link.X1, link.X2) >= line.XStart) return true;
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

    /// <summary>
    /// True when a word continues the previous one on the same line without a space, as the separate pieces
    /// around a ligature or a font change do.
    /// </summary>
    internal static bool ContinuesContentsWord(PdfUnderstandingWord previous, PdfUnderstandingWord word) =>
        Math.Abs(previous.BaselineY - word.BaselineY) <= Math.Max(1D, word.FontSize * 0.3D) &&
        word.XStart - previous.XEnd < Math.Max(previous.FontSize, word.FontSize) * 0.12D &&
        word.XStart >= previous.XStart;

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
        out bool hasLeader,
        out string? lastTitleText) {
        title = new List<PdfUnderstandingWord>();
        pageLabel = null;
        hasLeader = false;
        lastTitleText = null;
        int end = words.Count;
        if (end == 0) return false;
        PdfUnderstandingWord last = words[end - 1];
        string labelText = last.Text.Trim();
        if (TrySplitGluedLabel(labelText, strictPageLabel, out string? gluedLabel, out string? gluedTitle)) {
            // "Part C..........4": the title, leader and page number were painted without spaces.
            hasLeader = true;
            pageLabel = gluedLabel;
            title.AddRange(words);
            lastTitleText = gluedTitle;
            return words.Take(words.Count - 1).Any(static word => word.Text.Any(char.IsLetter)) || gluedTitle!.Any(char.IsLetter);
        }
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
                double gap = last.XStart - words[end - 2].XEnd;
                bool wideGap = gap >= Math.Max(last.FontSize, previous.FontSize) * 1.5D;
                // Dots that run up to the page number are its leader; dots followed by a wide gap end the title.
                bool trailingLeader = titleEnd == end - 1 && !wideGap && CountTrailing(previous.Text.Trim()) >= 3;
                if (leader || trailingLeader || wideGap) {
                    hasLeader = leader || trailingLeader;
                    pageLabel = candidateLabel;
                    end = titleEnd;
                    if (trailingLeader) {
                        string text = previous.Text.Trim();
                        lastTitleText = text.Substring(0, text.Length - CountTrailing(text)).TrimEnd();
                    }
                }
            }
        }
        if (strictPageLabel && pageLabel is null) return false;
        // Only the separator before the page number is a leader; dots or an ellipsis inside the title stay.
        if (lastTitleText is null) {
            while (end > 0 && IsLeaderWord(words[end - 1].Text)) end--;
        }
        for (int index = 0; index < end; index++) title.Add(words[index]);
        if (lastTitleText is { Length: 0 }) {
            title.RemoveAt(title.Count - 1);
            lastTitleText = null;
        }
        return title.Count > 0 && title.Any(static word => word.Text.Any(char.IsLetter));
    }

    /// <summary>
    /// Splits "Title.......12" painted as one word. A glued leader must be longer than an ellipsis, so a title such
    /// as "Release...2026" is never read as a page number.
    /// </summary>
    private static bool TrySplitGluedLabel(string text, bool strict, out string? label, out string? title) {
        label = null;
        title = null;
        int labelStart = text.Length;
        while (labelStart > 0 && !IsLeaderCharacter(text[labelStart - 1])) labelStart--;
        int leaderStart = labelStart;
        while (leaderStart > 0 && IsLeaderCharacter(text[leaderStart - 1])) leaderStart--;
        if (labelStart == text.Length || labelStart - leaderStart < 5 || leaderStart == 0) return false;
        string candidate = text.Substring(labelStart);
        if (!IsPageLabel(candidate, strict)) return false;
        label = candidate;
        title = text.Substring(0, leaderStart).TrimEnd();
        return true;
    }

    // Dot, hyphen and underscore leaders, plus the middle-dot and ellipsis forms other producers paint.
    private static bool IsLeaderCharacter(char ch) => ch is '.' or '-' or '_' or '\u00B7' or '\u2026' or '\u2022' or '\u2219' or '\u22C5';

    internal static bool IsContentsLeaderWord(string text) => IsLeaderWord(text);

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
        internal List<PdfUnderstandingWord> Words { get; } = new();

        internal void Add(PdfUnderstandingLine line) {
            Lines.Add(line);
            Words.AddRange(line.Words);
        }

        internal void Complete() {
            if (Lines.Count > 1) Words.Sort(static (left, right) => left.XStart.CompareTo(right.XStart));
        }
    }

    private sealed class PatternContentsLine {
        internal PatternContentsLine(ContentsRow row, int rowIndex, List<PdfUnderstandingWord> words, List<PdfUnderstandingWord> title,
            string pageLabel, bool hasLeader, string? lastTitleText) {
            LastTitleText = lastTitleText;
            Row = row;
            RowIndex = rowIndex;
            Words = words;
            Title = title;
            PageLabel = pageLabel;
            HasLeader = hasLeader;
        }

        internal ContentsRow Row { get; }
        internal int RowIndex { get; }
        internal List<PdfUnderstandingWord> Words { get; }
        internal List<PdfUnderstandingWord> Title { get; }
        internal string PageLabel { get; }
        internal bool HasLeader { get; }
        internal string? LastTitleText { get; }
        internal string DisplayTitle {
            get {
                var builder = new StringBuilder();
                for (int index = 0; index < Title.Count; index++) {
                    if (index > 0 && !ContinuesContentsWord(Title[index - 1], Title[index])) builder.Append(' ');
                    builder.Append(index == Title.Count - 1 && LastTitleText is not null ? LastTitleText : Title[index].Text.Trim());
                }
                return builder.ToString();
            }
        }
        // Indentation measured from the right-aligned page numbers, so columns of equal width share levels.
        internal double Indent => XEnd - Title[0].XStart;
        internal double XStart => Words[0].XStart;
        internal double XEnd => Words[Words.Count - 1].XEnd;
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
        PdfInferenceEvidence evidence,
        string? lastTitleText) {
        Level = level;
        LastTitleText = lastTitleText;
        // The words that show the page number; a glued "Title.....4" shows it in the last title word.
        PdfUnderstandingWord[] labelWords = pageLabel is null
            ? Array.Empty<PdfUnderstandingWord>()
            : words.Where(word => !title.Contains(word) && !PdfDocumentSemanticEnricher.IsContentsLeaderWord(word.Text)).ToArray();
        PageLabelWords = labelWords.Length == 0 && pageLabel is not null && title.Count > 0 ? new[] { title[title.Count - 1] } : labelWords;
        Title = title;
        PageLabel = pageLabel;
        Lines = lines;
        Evidence = evidence;
        Runs = new HashSet<PdfTextSpan>(words.SelectMany(static word => word.SourceRuns));
    }

    internal int Level { get; }
    internal IReadOnlyList<PdfUnderstandingWord> Title { get; }
    /// <summary>The last title word without the leader or page number glued to it, when the parser removed one.</summary>
    internal string? LastTitleText { get; }
    internal IReadOnlyList<PdfUnderstandingWord> PageLabelWords { get; }
    internal string? PageLabel { get; }
    internal IReadOnlyList<PdfUnderstandingLine> Lines { get; }
    internal HashSet<PdfTextSpan> Runs { get; }
    internal PdfInferenceEvidence Evidence { get; }
}

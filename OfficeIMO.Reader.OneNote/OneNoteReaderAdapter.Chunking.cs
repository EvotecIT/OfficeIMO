using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Markdown;

namespace OfficeIMO.Reader.OneNote;

internal static partial class OneNoteReaderAdapter {
    private static IReadOnlyList<ProjectionPart> BuildProjectionParts(OneNotePage page, int pageIndex, int maxChars) {
        maxChars = Math.Max(1, maxChars);
        IReadOnlyDictionary<OneNoteBinaryElement, string> assetTargets = BuildAssetTargets(page, pageIndex);
        Func<OneNoteBinaryElement, string?> resolver = element =>
            assetTargets.TryGetValue(element, out string? target) ? target : null;
        var units = new List<ProjectionPart>();

        var heading = new OneNotePage { Title = page.Title, Level = page.Level };
        int headingLevel = Math.Min(5, Math.Max(0, page.Level)) + 1;
        units.Add(new ProjectionPart(
            OneNoteMarkdownProjection.ToText(heading),
            OneNoteMarkdownProjection.ToMarkdown(heading, headingLevel, resolver)));

        foreach (OneNoteOutline outline in page.Outlines) {
            foreach (OneNoteElement child in outline.Children) AddElementUnits(child, resolver, maxChars, units);
        }
        foreach (OneNoteElement element in page.DirectContent) AddElementUnits(element, resolver, maxChars, units);

        return PackUnits(units, maxChars);
    }

    private static IReadOnlyDictionary<OneNoteBinaryElement, string> BuildAssetTargets(OneNotePage page, int pageIndex) {
        var targets = new Dictionary<OneNoteBinaryElement, string>();
        int assetIndex = 0;
        foreach (OneNoteElement element in EnumerateAllElements(page)) {
            if (!(element is OneNoteBinaryElement binary)) continue;
            string assetId = BuildAssetId(pageIndex, assetIndex++);
            if (binary.Payload != null) targets[binary] = assetId;
        }
        return targets;
    }

    private static void AddElementUnits(
        OneNoteElement element,
        Func<OneNoteBinaryElement, string?> resolver,
        int maxChars,
        ICollection<ProjectionPart> units) {
        if (element is OneNoteOutline outline) {
            foreach (OneNoteElement child in outline.Children) AddElementUnits(child, resolver, maxChars, units);
            return;
        }

        ProjectionPart whole = ProjectElement(element, resolver);
        if (whole.Fits(maxChars)) {
            units.Add(whole);
            return;
        }

        if (element is OneNoteParagraph paragraph) {
            bool firstRun = true;
            foreach (OneNoteTextRun run in paragraph.Runs) {
                AddRunUnits(run, firstRun ? paragraph.List : null, resolver, maxChars, units);
                firstRun = false;
            }
            foreach (OneNoteElement child in paragraph.Children) AddElementUnits(child, resolver, maxChars, units);
            return;
        }

        if (element is OneNoteTable table) {
            foreach (OneNoteTableRow row in table.Rows) {
                var rowTable = new OneNoteTable();
                foreach (double width in table.ColumnWidths) rowTable.ColumnWidths.Add(width);
                rowTable.Rows.Add(row);
                ProjectionPart rowProjection = ProjectElement(rowTable, resolver);
                if (rowProjection.Fits(maxChars)) {
                    units.Add(rowProjection);
                } else {
                    foreach (OneNoteTableCell cell in row.Cells) {
                        foreach (OneNoteElement cellElement in cell.Content) {
                            AddElementUnits(cellElement, resolver, maxChars, units);
                        }
                    }
                }
            }
            return;
        }

        units.Add(whole);
    }

    private static void AddRunUnits(
        OneNoteTextRun run,
        OneNoteListInfo? list,
        Func<OneNoteBinaryElement, string?> resolver,
        int maxChars,
        ICollection<ProjectionPart> units) {
        int offset = 0;
        bool firstSegment = true;
        while (offset < run.Text.Length) {
            int length = FindRunSegmentLength(run, list, resolver, maxChars, offset, firstSegment);
            if (length > 0 && offset + length < run.Text.Length) {
                int boundary = FindWhitespaceBoundary(run.Text, offset, length);
                if (boundary > offset) length = boundary - offset;
            }
            length = KeepSurrogatePairTogether(run.Text, offset, length);
            if (length <= 0) {
                length = ScalarLengthAt(run.Text, offset);
                units.Add(ProjectRun(run, run.Text.Substring(offset, length), firstSegment ? list : null, resolver));
                offset += length;
            } else {
                units.Add(ProjectRun(run, run.Text.Substring(offset, length), firstSegment ? list : null, resolver));
                offset += length;
            }
            firstSegment = false;
        }

        if (run.Text.Length == 0) units.Add(ProjectRun(run, string.Empty, list, resolver));
    }

    private static int KeepSurrogatePairTogether(string value, int offset, int length) {
        int end = offset + length;
        return length > 0 && end < value.Length &&
               char.IsHighSurrogate(value[end - 1]) && char.IsLowSurrogate(value[end])
            ? length - 1
            : length;
    }

    private static int ScalarLengthAt(string value, int offset) =>
        char.IsHighSurrogate(value[offset]) && offset + 1 < value.Length && char.IsLowSurrogate(value[offset + 1]) ? 2 : 1;

    private static int FindWhitespaceBoundary(string value, int offset, int length) {
        int whitespace = offset + length - 1;
        while (whitespace >= offset && !char.IsWhiteSpace(value[whitespace])) whitespace--;
        if (whitespace <= offset) return offset;
        while (whitespace > offset && char.IsWhiteSpace(value[whitespace - 1])) whitespace--;
        return whitespace;
    }

    private static int FindRunSegmentLength(
        OneNoteTextRun run,
        OneNoteListInfo? list,
        Func<OneNoteBinaryElement, string?> resolver,
        int maxChars,
        int offset,
        bool firstSegment) {
        int low = 1;
        int high = Math.Min(maxChars, run.Text.Length - offset);
        int best = 0;
        while (low <= high) {
            int length = low + ((high - low) / 2);
            ProjectionPart candidate = ProjectRun(
                run,
                run.Text.Substring(offset, length),
                firstSegment ? list : null,
                resolver);
            if (candidate.Fits(maxChars)) {
                best = length;
                low = length + 1;
            } else {
                high = length - 1;
            }
        }
        return best;
    }

    private static ProjectionPart ProjectRun(
        OneNoteTextRun source,
        string text,
        OneNoteListInfo? list,
        Func<OneNoteBinaryElement, string?> resolver) {
        text = OneNoteTextProjection.Normalize(text);
        var paragraph = new OneNoteParagraph { List = list };
        var run = new OneNoteTextRun {
            Text = text,
            Hyperlink = source.Hyperlink,
            HyperlinkProtected = source.HyperlinkProtected
        };
        run.Style.FontFamily = source.Style.FontFamily;
        run.Style.FontSize = source.Style.FontSize;
        run.Style.ColorArgb = source.Style.ColorArgb;
        run.Style.HighlightColorArgb = source.Style.HighlightColorArgb;
        run.Style.Bold = source.Style.Bold;
        run.Style.Italic = source.Style.Italic;
        run.Style.Underline = source.Style.Underline;
        run.Style.Strikethrough = source.Style.Strikethrough;
        run.Style.Superscript = source.Style.Superscript;
        run.Style.Subscript = source.Style.Subscript;
        run.Style.LanguageId = source.Style.LanguageId;
        run.Style.IsMath = source.Style.IsMath;
        paragraph.Runs.Add(run);
        ProjectionPart projected = ProjectElement(paragraph, resolver);
        string plainText = (list == null ? string.Empty : list.Ordered ? "1. " : "- ") + text;
        string markdown = string.IsNullOrWhiteSpace(text) ? plainText : projected.Markdown;
        return new ProjectionPart(plainText, markdown);
    }

    private static ProjectionPart ProjectElement(
        OneNoteElement element,
        Func<OneNoteBinaryElement, string?> resolver) {
        return new ProjectionPart(
            OneNoteMarkdownProjection.ToText(element),
            OneNoteMarkdownProjection.ToMarkdown(element, resolver));
    }

    private static IReadOnlyList<ProjectionPart> PackUnits(IReadOnlyList<ProjectionPart> units, int maxChars) {
        var result = new List<ProjectionPart>();
        var text = new StringBuilder();
        var markdown = new StringBuilder();

        foreach (ProjectionPart source in units) {
            string textSeparator = text.Length == 0 || source.Text.Length == 0 ? string.Empty : Environment.NewLine;
            string markdownSeparator = markdown.Length == 0 || source.Markdown.Length == 0
                ? string.Empty
                : Environment.NewLine + Environment.NewLine;
            bool fits = text.Length + textSeparator.Length + source.Text.Length <= maxChars &&
                        markdown.Length + markdownSeparator.Length + source.Markdown.Length <= maxChars;
            if (!fits && (text.Length > 0 || markdown.Length > 0)) {
                result.Add(new ProjectionPart(text.ToString(), markdown.ToString()));
                text.Clear();
                markdown.Clear();
                textSeparator = string.Empty;
                markdownSeparator = string.Empty;
            }
            text.Append(textSeparator).Append(source.Text);
            markdown.Append(markdownSeparator).Append(source.Markdown);
        }

        if (text.Length > 0 || markdown.Length > 0 || result.Count == 0) {
            result.Add(new ProjectionPart(text.ToString(), markdown.ToString()));
        }
        return result;
    }

    private readonly struct ProjectionPart {
        internal ProjectionPart(string text, string markdown) {
            Text = text;
            Markdown = markdown;
        }

        internal string Text { get; }
        internal string Markdown { get; }
        internal bool Fits(int maxChars) => Text.Length <= maxChars && Markdown.Length <= maxChars;
    }
}

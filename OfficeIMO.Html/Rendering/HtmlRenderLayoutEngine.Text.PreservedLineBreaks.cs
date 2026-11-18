namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static readonly char[] SegmentBreakCharacters = { '\n', '\r' };

    /// <summary>Text runs for anonymous text that bypasses DOM run collection, such as anonymous flex items.</summary>
    private static IReadOnlyList<HtmlInlineRun> CreateAnonymousTextRuns(string text, HtmlRenderBoxStyle style, string? link, string source) {
        if (style.CollapseSegmentBreaks) text = ReplaceSegmentBreaks(text);
        if (!style.PreserveLineBreaks || style.PreserveWhitespace || text.IndexOfAny(SegmentBreakCharacters) < 0)
            return new[] { new HtmlInlineRun(text, style, link, source) };
        var runs = new List<HtmlInlineRun>();
        AddPreservedLineBreakRuns(text, value => runs.Add(new HtmlInlineRun(value, style, link, source)),
            () => runs.Add(new HtmlInlineRun("\u2028", style, link, source)));
        return runs;
    }

    /// <summary>white-space-collapse: preserve-spaces turns each segment break into a space.</summary>
    private static string ReplaceSegmentBreaks(string text) =>
        text.IndexOfAny(SegmentBreakCharacters) < 0 ? text : text.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ');

    /// <summary>
    /// Splits <c>white-space: pre-line</c> text at its segment breaks. Each break becomes a forced
    /// line break, and the collapsible spaces and tabs around it are removed, as CSS Text requires.
    /// </summary>
    private static void AddPreservedLineBreakRuns(string text, Action<string> addText, Action addBreak) {
        int start = 0;
        for (int index = 0; index < text.Length; index++) {
            char value = text[index];
            if (value != '\n' && value != '\r') continue;
            AddSegment(text.Substring(start, index - start), trimStart: start > 0, trimEnd: true);
            addBreak();
            if (value == '\r' && index + 1 < text.Length && text[index + 1] == '\n') index++;
            start = index + 1;
        }
        AddSegment(text.Substring(start), trimStart: start > 0, trimEnd: false);

        void AddSegment(string segment, bool trimStart, bool trimEnd) {
            if (trimStart) segment = segment.TrimStart(' ', '\t');
            if (trimEnd) segment = segment.TrimEnd(' ', '\t');
            if (segment.Length > 0) addText(segment);
        }
    }
}

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static readonly char[] SegmentBreakCharacters = { '\n', '\r' };

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

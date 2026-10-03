namespace OfficeIMO.Reader.Latex;

// Split the shared typed text projection, then encode each fragment according to its
// provenance. Literal text cannot become Markdown syntax, and only actual labels are HTML.
internal static class LatexReaderTextSplitter {
    internal static IReadOnlyList<LatexReaderPart> Split(IReadOnlyList<LatexTextSegment> segments,
        int maximum, CancellationToken cancellationToken) {
        string text = LatexProjectedText.Text(segments);
        if (text.Length == 0) return Array.Empty<LatexReaderPart>();
        if (maximum <= 0 || text.Length <= maximum) return new[] { new LatexReaderPart(text, string.Empty) };
        var ends = new int[segments.Count];
        int position = 0;
        for (int index = 0; index < segments.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            position += segments[index].Text.Length;
            ends[index] = position;
        }
        var parts = new List<LatexReaderPart>();
        int offset = 0, segmentIndex = 0;
        while (offset < text.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            while (segmentIndex < ends.Length && ends[segmentIndex] <= offset) segmentIndex++;
            int end = offset + Math.Min(maximum, text.Length - offset);
            if (end < text.Length) {
                int split = text.LastIndexOf('\n', end - 1, end - offset);
                if (split < offset) split = text.LastIndexOf(' ', end - 1, end - offset);
                if (split >= offset) end = split + 1;
                if (end < text.Length && char.IsHighSurrogate(text[end - 1]) && char.IsLowSurrogate(text[end]))
                    end = end - offset == 1 ? end + 1 : end - 1;
            }
            for (int index = segmentIndex; index < segments.Count && SegmentStart(ends, index) < end; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (segments[index].IsAnchor && end < ends[index]) {
                    int start = SegmentStart(ends, index);
                    end = start > offset ? start : ends[index];
                    break;
                }
            }
            string markdown = Render(segments, ends, segmentIndex, offset, end, cancellationToken);
            parts.Add(new LatexReaderPart(text.Substring(offset, end - offset), markdown));
            offset = end;
        }
        return parts;
    }

    private static string Render(IReadOnlyList<LatexTextSegment> segments, int[] ends, int first,
        int start, int end, CancellationToken cancellationToken) {
        var markdown = new StringBuilder();
        var plain = new StringBuilder();
        for (int index = first; index < segments.Count && SegmentStart(ends, index) < end; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            LatexTextSegment segment = segments[index];
            int segmentStart = SegmentStart(ends, index);
            int from = Math.Max(start, segmentStart), to = Math.Min(end, ends[index]);
            if (to <= from) continue;
            string fragment = segment.Text.Substring(from - segmentStart, to - from);
            if (segment.IsAnchor || segment.Language != null) {
                FlushPlain(markdown, plain);
                markdown.Append("\n\n");
                markdown.Append(segment.IsAnchor ? fragment : MarkdownDoc.Create().Add(new CodeBlock(segment.Language!, fragment)).ToMarkdown().TrimEnd());
                markdown.Append("\n\n");
            } else plain.Append(fragment);
        }
        FlushPlain(markdown, plain);
        return markdown.ToString().Trim('\r', '\n');
    }

    private static void FlushPlain(StringBuilder markdown, StringBuilder plain) {
        if (plain.Length == 0) return;
        markdown.Append(MarkdownEscaper.EscapeLiteralText(plain.ToString()));
        plain.Clear();
    }
    private static int SegmentStart(int[] ends, int index) => index == 0 ? 0 : ends[index - 1];
}

internal sealed class LatexReaderPart {
    internal LatexReaderPart(string text, string markdown) { Text = text; Markdown = markdown; }
    internal string Text { get; }
    internal string Markdown { get; }
}

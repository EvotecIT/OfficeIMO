namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookScriptImporter {
    private static (char Marker, int Length, bool CanClose) ReadFence(ReadOnlySpan<char> text) {
        int indent = 0;
        while (indent < text.Length && text[indent] == ' ' && indent < 4) indent++;
        if (indent == 4) return default;
        text = text.Slice(indent);
        if (text.Length < 3 || text[0] is not ('`' or '~')) return default;
        int length = 1;
        while (length < text.Length && text[length] == text[0]) length++;
        return length >= 3 ? (text[0], length, text.Slice(length).Trim().IsEmpty) : default;
    }

    // A percent marker inside commented Markdown is a boundary unless the enclosing fence has a closer
    // before the next uncommented line of source. Read each run once, even when it has unmatched openers.
    private sealed class PercentFenceIndex(StringReader reader, CancellationToken cancellationToken) {
        private (int Length, int LastLine)[] _backticks = [];
        private (int Length, int LastLine)[] _tildes = [];
        private int _lineNumber;

        internal bool HasCloser(char marker, int length, int line) {
            if (line >= _lineNumber) ReadCommentRun(line);
            (int Length, int LastLine)[] entries = marker == '`' ? _backticks : _tildes;
            int lower = 0;
            int upper = entries.Length;
            while (lower < upper) {
                int middle = lower + (upper - lower) / 2;
                if (entries[middle].Length < length) lower = middle + 1;
                else upper = middle;
            }
            return lower < entries.Length && entries[lower].LastLine > line;
        }

        private void ReadCommentRun(int openingLine) {
            var backticks = new Dictionary<int, int>();
            var tildes = new Dictionary<int, int>();
            while (reader.ReadLine() is { } line) {
                cancellationToken.ThrowIfCancellationRequested();
                if (++_lineNumber <= openingLine || string.IsNullOrWhiteSpace(line)) continue;
                ReadOnlySpan<char> text = RemoveComment(line.AsSpan());
                if (text.Length == line.Length) break;
                (char marker, int length, bool canClose) = ReadFence(text);
                if (canClose) (marker == '`' ? backticks : tildes)[length] = _lineNumber;
            }
            _backticks = LastClosingLines(backticks);
            _tildes = LastClosingLines(tildes);
        }

        private static (int Length, int LastLine)[] LastClosingLines(Dictionary<int, int> closers) {
            var entries = closers.OrderBy(entry => entry.Key).Select(entry => (Length: entry.Key, LastLine: entry.Value)).ToArray();
            for (int i = entries.Length - 2; i >= 0; i--) entries[i].LastLine = Math.Max(entries[i].LastLine, entries[i + 1].LastLine);
            return entries;
        }
    }
}

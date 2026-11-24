namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookScriptImporter {
    private static (char Marker, int Length, bool CanClose) ReadFence(ReadOnlySpan<char> text) {
        text = text.TrimStart();
        if (text.Length < 3 || text[0] is not ('`' or '~')) return default;
        int length = 1;
        while (length < text.Length && text[length] == text[0]) length++;
        return length >= 3 ? (text[0], length, text.Slice(length).Trim().IsEmpty) : default;
    }

    // A percent marker inside Markdown is a cell boundary unless the enclosing fence has a closer.
    // Build a suffix lookup only for scripts containing fenced Markdown, so unmatched openers cannot make repeated scans.
    private sealed class PercentFenceIndex(string script, CancellationToken cancellationToken) {
        private (int Length, int LastLine)[]? _backticks;
        private (int Length, int LastLine)[]? _tildes;

        internal bool HasCloser(char marker, int length, int line) {
            if (_backticks == null || _tildes == null) Build();
            (int Length, int LastLine)[] entries = marker == '`' ? _backticks! : _tildes!;
            int lower = 0;
            int upper = entries.Length;
            while (lower < upper) {
                int middle = lower + (upper - lower) / 2;
                if (entries[middle].Length < length) lower = middle + 1;
                else upper = middle;
            }
            return lower < entries.Length && entries[lower].LastLine > line;
        }

        private void Build() {
            var backticks = new Dictionary<int, int>();
            var tildes = new Dictionary<int, int>();
            using var reader = new StringReader(script);
            int lineNumber = 0;
            while (reader.ReadLine() is { } line) {
                cancellationToken.ThrowIfCancellationRequested();
                lineNumber++;
                (char marker, int length, bool canClose) = ReadFence(RemoveComment(line.AsSpan()));
                if (canClose) (marker == '`' ? backticks : tildes)[length] = lineNumber;
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

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookMarkupImporter {
    private readonly record struct Fence(char Marker, int Length, int InfoOffset);

    private static bool TryOpenFence(ReadOnlySpan<char> line, out Fence fence) {
        fence = default;
        int index = LeadingSpaces(line);
        if (index > 3 || index == line.Length || line[index] is not ('`' or '~')) return false;
        char marker = line[index];
        int start = index;
        while (index < line.Length && line[index] == marker) index++;
        int length = index - start;
        if (length < 3 || (marker == '`' && line.Slice(index).IndexOf('`') >= 0)) return false;
        fence = new Fence(marker, length, index);
        return true;
    }

    private static bool IsClosingFence(ReadOnlySpan<char> line, Fence fence) {
        int start = LeadingSpaces(line);
        if (start > 3 || line.Length - start < fence.Length) return false;
        int end = start;
        while (end < line.Length && line[end] == fence.Marker) end++;
        return end - start >= fence.Length && line.Slice(end).Trim().IsEmpty;
    }

    private static int LeadingSpaces(ReadOnlySpan<char> line) {
        int index = 0;
        while (index < line.Length && line[index] == ' ') index++;
        return index;
    }

    private static bool TryGetLanguage(ReadOnlySpan<char> info, NotebookImportFormat format, string? fallback, out string? language) {
        language = null;
        if (format == NotebookImportFormat.MySTMarkdown) {
            ReadOnlySpan<char> directive = "{code-cell}".AsSpan();
            if (!info.StartsWith(directive, StringComparison.Ordinal)) return false;
            ReadOnlySpan<char> argument = info.Slice(directive.Length);
            if (!argument.IsEmpty && !char.IsWhiteSpace(argument[0])) return false;
            argument = argument.Trim();
            if (argument.IsEmpty) {
                language = fallback;
                return true;
            }
            return TryReadLanguage(argument, out language);
        }

        if (info.Length < 3 || info[0] != '{' || info[info.Length - 1] != '}') return false;
        ReadOnlySpan<char> arguments = info.Slice(1, info.Length - 2).TrimStart();
        return TryReadLanguage(arguments, out language);
    }

    private static bool TryReadLanguage(ReadOnlySpan<char> value, out string? language) {
        language = null;
        if (value.IsEmpty || !char.IsLetter(value[0])) return false;
        int end = 0;
        while (end < value.Length && (char.IsLetterOrDigit(value[end]) || value[end] is '_' or '-' or '+' or '#' or '.')) end++;
        if (end > 64 || (end < value.Length && !char.IsWhiteSpace(value[end]) && value[end] != ',')) return false;
        language = value.Slice(0, end).ToString();
        return true;
    }

    private static bool IsBracedDirective(ReadOnlySpan<char> info, NotebookImportFormat format) =>
        !info.IsEmpty && info[0] == '{' && (format == NotebookImportFormat.MySTMarkdown
            || info.Length > 1 && info[1] is '=');

    private static int ReadCellOptions(string source, int start, int end, NotebookImportFormat format, CancellationToken token) {
        var cursor = new SourceCursor(source, start, end);
        if (!cursor.TryRead(out SourceLine first)) return start;
        ReadOnlySpan<char> text = first.Text(source).TrimStart();
        if (format == NotebookImportFormat.MySTMarkdown && text.SequenceEqual("---".AsSpan())) {
            while (cursor.TryRead(out SourceLine option)) {
                token.ThrowIfCancellationRequested();
                if (option.Text(source).Trim().SequenceEqual("---".AsSpan())) return cursor.Offset;
            }
            return start;
        }

        int endOfOptions = start;
        SourceLine line = first;
        do {
            token.ThrowIfCancellationRequested();
            text = line.Text(source).TrimStart();
            if (format == NotebookImportFormat.MySTMarkdown) {
                if (!IsMySTOption(text)) break;
            } else if (!text.StartsWith("#|".AsSpan(), StringComparison.Ordinal)) {
                break;
            }
            endOfOptions = cursor.Offset;
        } while (cursor.TryRead(out line));
        return endOfOptions;
    }

    private static bool IsMySTOption(ReadOnlySpan<char> text) {
        if (text.Length < 3 || text[0] != ':') return false;
        int delimiter = text.Slice(1).IndexOf(':') + 1;
        return delimiter > 1 && text.Slice(1, delimiter - 1).IndexOfAny(' ', '\t') < 0;
    }

    private readonly record struct SourceLine(int Start, int ContentLength, int Number) {
        internal ReadOnlySpan<char> Text(string source) => source.AsSpan(Start, ContentLength);
    }

    private struct SourceCursor {
        private readonly int _end;
        private int _lineNumber;

        internal SourceCursor(string source, int start = 0, int? end = null) {
            Source = source;
            Offset = start;
            _end = end ?? source.Length;
            _lineNumber = 0;
        }

        internal string Source { get; }
        internal int Offset { get; private set; }

        internal bool TryRead(out SourceLine line) {
            line = default;
            int start = Offset;
            if (start == _end) return false;
            int end = start;
            while (end < _end && Source[end] is not ('\r' or '\n')) end++;
            line = new SourceLine(start, end - start, ++_lineNumber);
            if (end < _end && Source[end] == '\r') end++;
            if (end < _end && Source[end] == '\n') end++;
            Offset = end;
            return true;
        }
    }
}

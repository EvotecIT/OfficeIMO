namespace OfficeIMO.Markdown;

/// <summary>
/// The lines of one block parse, kept as offsets into the source text. A line becomes a string
/// only when a parser asks for it as one, and that string is kept for later reads.
/// </summary>
internal sealed class MarkdownSourceLines : IReadOnlyList<string> {
    private static readonly char[] LineEndingCharacters = { '\r', '\n' };

    private readonly string _text;
    // Start offset in _text of each line; a line ends before the line ending in front of the next
    // start. Null when every line came in as a string.
    private readonly int[]? _starts;
    // Whether _text has CR line endings, so a line end has to look at the characters.
    private readonly bool _hasCarriageReturn;
    private string?[]? _strings;
    // Set once a line may differ from the source text: a line was replaced, or the string array
    // was handed out and may be written to.
    private bool _mayDifferFromText;
    // Set once every line has a string, so ToArray hands the array out without checking it again.
    private bool _allLinesCreated;
    // The last FindBlankOrStartingWith answer: from _probeStart the first matching line is _probeResult.
    private int _probeStart = -1;
    private int _probeResult;
    private char _probeCharacter;

    private MarkdownSourceLines(string text, int[] starts, bool hasCarriageReturn) {
        _text = text;
        _starts = starts;
        _hasCarriageReturn = hasCarriageReturn;
        Length = starts.Length;
    }

    private MarkdownSourceLines(string[] lines) {
        _text = string.Empty;
        _strings = lines;
        Length = lines.Length;
    }

    /// <summary>Number of lines.</summary>
    public int Length { get; }

    int IReadOnlyCollection<string>.Count => Length;

    /// <summary>
    /// The line as a string. Reading it creates the string once; writing it replaces the line for
    /// every later read, slice reads included.
    /// </summary>
    public string this[int index] {
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        get {
            var strings = _strings;
            if (strings != null) {
                var line = strings[index];
                if (line != null) {
                    return line;
                }
            }
            return GetOrCreateLine(index);
        }
        set {
            if ((uint)index >= (uint)Length) {
                throw new IndexOutOfRangeException();
            }
            if (value == null && !_allLinesCreated) {
                // A null slot reads as "not created yet"; create the rest first so this null stays a null.
                ToArray();
            }
            (_strings ??= new string?[Length])[index] = value;
            _mayDifferFromText = true;
        }
    }

    private string GetOrCreateLine(int index) {
        var strings = _strings;
        if (_starts == null || _allLinesCreated) {
            // Lines handed in as strings are returned as they are, null included; once every line has
            // been created, a null was written by a parser and stays one.
            return strings![index]!;
        }
        if ((uint)index >= (uint)Length) {
            throw new IndexOutOfRangeException();
        }

        strings ??= _strings = new string?[Length];
        int start = _starts[index];
        var created = _text.Substring(start, GetEnd(index) - start);
        strings[index] = created;
        return created;
    }

    /// <summary>Wraps lines that already exist as strings; writes through the line source change the array.</summary>
    public static implicit operator MarkdownSourceLines(string[] lines) => FromArray(lines);

    /// <summary>Wraps lines that already exist as strings.</summary>
    public static MarkdownSourceLines FromArray(string[] lines) => new(lines ?? Array.Empty<string>());

    /// <summary>
    /// Indexes the lines of <paramref name="text"/>, split at CRLF, CR and LF, without copying them.
    /// </summary>
    public static MarkdownSourceLines FromText(string text) {
        text ??= string.Empty;
        bool hasCarriageReturn = text.IndexOf('\r') >= 0;
        int lineCount = 1;
        if (hasCarriageReturn) {
            for (int i = text.IndexOfAny(LineEndingCharacters); i >= 0; i = text.IndexOfAny(LineEndingCharacters, i + 1)) {
                lineCount++;
                if (text[i] == '\r' && i + 1 < text.Length && text[i + 1] == '\n') {
                    i++;
                }
            }
        } else {
#if NET8_0_OR_GREATER
            lineCount += text.AsSpan().Count('\n');
#else
            for (int i = text.IndexOf('\n'); i >= 0; i = text.IndexOf('\n', i + 1)) {
                lineCount++;
            }
#endif
        }

        var starts = new int[lineCount];
#if NET8_0_OR_GREATER
        if (!hasCarriageReturn) {
            FillLineFeedStarts(text, starts);
            return new MarkdownSourceLines(text, starts, hasCarriageReturn: false);
        }
#endif
        int lineIndex = 1;
        for (int i = hasCarriageReturn ? text.IndexOfAny(LineEndingCharacters) : text.IndexOf('\n');
             i >= 0;
             i = hasCarriageReturn ? text.IndexOfAny(LineEndingCharacters, i + 1) : text.IndexOf('\n', i + 1)) {
            if (text[i] == '\r' && i + 1 < text.Length && text[i + 1] == '\n') {
                i++;
            }
            starts[lineIndex++] = i + 1;
        }

        return new MarkdownSourceLines(text, starts, hasCarriageReturn);
    }

#if NET8_0_OR_GREATER
    // One pass over text that has only LF line endings, eight characters at a time: lines here are
    // short, so searching for each line ending separately costs more than reading the text once.
    private static void FillLineFeedStarts(string text, int[] starts) {
        int lineIndex = 1;
        int index = 0;
        if (System.Runtime.Intrinsics.Vector128.IsHardwareAccelerated) {
            ref ushort first = ref System.Runtime.CompilerServices.Unsafe.As<char, ushort>(
                ref System.Runtime.InteropServices.MemoryMarshal.GetReference(text.AsSpan()));
            var lineFeed = System.Runtime.Intrinsics.Vector128.Create((ushort)'\n');
            int lastBlock = text.Length - System.Runtime.Intrinsics.Vector128<ushort>.Count;
            for (; index <= lastBlock; index += System.Runtime.Intrinsics.Vector128<ushort>.Count) {
                var block = System.Runtime.Intrinsics.Vector128.LoadUnsafe(ref first, (nuint)index);
                uint matches = System.Runtime.Intrinsics.Vector128.ExtractMostSignificantBits(System.Runtime.Intrinsics.Vector128.Equals(block, lineFeed));
                while (matches != 0) {
                    starts[lineIndex++] = index + System.Numerics.BitOperations.TrailingZeroCount(matches) + 1;
                    matches &= matches - 1;
                }
            }
        }

        for (; index < text.Length; index++) {
            if (text[index] == '\n') {
                starts[lineIndex++] = index + 1;
            }
        }
    }
#endif

    /// <summary>
    /// False only when no line can contain <paramref name="value"/>: every line still reads from the
    /// source text and the text does not contain it. Answers in one search instead of one per line.
    /// </summary>
    public bool MayContain(string value) =>
        _starts == null || _mayDifferFromText || _text.IndexOf(value, StringComparison.Ordinal) >= 0;

    /// <summary>
    /// Every line as a string. The returned array is the one this instance reads from, so writes
    /// to it are seen by later reads.
    /// </summary>
    public string[] ToArray() {
        _mayDifferFromText = true;
        if (_starts == null) {
            return (string[])_strings!;
        }

        var strings = _strings ??= new string?[Length];
        if (!_allLinesCreated) {
            for (int i = 0; i < strings.Length; i++) {
                if (strings[i] == null) {
                    _ = this[i];
                }
            }
            _allLinesCreated = true;
        }
        return (string[])strings!;
    }

    /// <summary>
    /// The line as a range of a string, without creating a string for it: either the source text
    /// or the string the line was replaced with.
    /// </summary>
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    public void GetSlice(int index, out string text, out int start, out int end) {
        var strings = _strings;
        if (strings != null) {
            var line = strings[index];
            if (line != null || _starts == null || _allLinesCreated) {
                // Lines handed in as strings, or written as null by a parser, read null as an empty line.
                text = line ?? string.Empty;
                start = 0;
                end = text.Length;
                return;
            }
        }

        text = _text;
        start = _starts![index];
        end = GetEnd(index);
    }

    // Where line index ends in _text, before its LF, CR or CRLF.
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private int GetEnd(int index) {
        int next = index + 1;
        if (next >= Length) {
            return _text.Length;
        }

        int end = _starts![next] - 1;
        if (_hasCarriageReturn && _text[end] == '\n' && end > _starts[index] && _text[end - 1] == '\r') {
            end--;
        }
        return end;
    }

    /// <summary>Same result as <c>CountLeadingIndentColumns</c> on the line: tabs advance to the next multiple of four.</summary>
    public int GetIndentColumns(int index) {
        GetSlice(index, out var text, out int start, out int end);
        int columns = 0;
        for (int i = start; i < end; i++) {
            char ch = text[i];
            if (ch == ' ') {
                columns++;
            } else if (ch == '\t') {
                columns += 4 - (columns % 4);
            } else {
                break;
            }
        }
        return columns;
    }

    /// <summary>
    /// False when the line cannot be a list item: its first non-white-space character is not a bullet,
    /// a digit or an ASCII letter (letter markers need list extras). Lets callers skip creating the line.
    /// </summary>
    public bool MayStartListItem(int index) {
        char first = GetFirstNonWhiteSpaceCharacter(index);
        // Digits as char.IsDigit reads them, as the ordered-list marker parser does.
        return first is '-' or '*' or '+' or (>= 'a' and <= 'z') or (>= 'A' and <= 'Z') || char.IsDigit(first);
    }

    /// <summary>
    /// The first line from <paramref name="start"/> that is blank or whose first non-white-space character is
    /// <paramref name="value"/>, or <see cref="Length"/> when there is none. The last answer is kept, so probing
    /// again from any line it covers costs nothing.
    /// </summary>
    public int FindBlankOrStartingWith(int start, char value) {
        if (!_mayDifferFromText && value == _probeCharacter && _probeStart >= 0 && start >= _probeStart && start <= _probeResult) {
            return _probeResult;
        }

        int index = start;
        while (index < Length && !IsBlank(index) && GetFirstNonWhiteSpaceCharacter(index) != value) {
            index++;
        }

        _probeCharacter = value;
        _probeStart = start;
        _probeResult = index;
        return index;
    }


    /// <summary>Whether the line contains <paramref name="value"/>, searched in place.</summary>
    public bool LineContains(int index, char value) {
        GetSlice(index, out var text, out int start, out int end);
        return text.IndexOf(value, start, end - start) >= 0;
    }

    /// <summary>Same result as <see cref="string.IsNullOrWhiteSpace(string)"/> on the line.</summary>
    public bool IsBlank(int index) {
        GetSlice(index, out var text, out int start, out int end);
        for (int i = start; i < end; i++) {
            if (!char.IsWhiteSpace(text[i])) {
                return false;
            }
        }
        return true;
    }

    /// <summary>Same result as <c>GetFirstNonWhiteSpaceCharacter</c> on the line.</summary>
    public char GetFirstNonWhiteSpaceCharacter(int index) {
        GetSlice(index, out var text, out int start, out int end);
        for (int i = start; i < end; i++) {
            char value = text[i];
            if (!char.IsWhiteSpace(value)) {
                return value;
            }
        }
        return '\0';
    }

    /// <summary>
    /// Lines <paramref name="start"/> to <paramref name="start"/> + <paramref name="count"/> - 1 joined with
    /// <paramref name="separator"/>, like <see cref="string.Join(string, string[], int, int)"/>.
    /// </summary>
    public string Join(string separator, int start, int count) {
        if (start < 0 || count < 0 || start + count > Length) {
            throw new ArgumentOutOfRangeException(nameof(count));
        }
        if (count == 0) {
            return string.Empty;
        }
        if (count == 1) {
            return this[start] ?? string.Empty;
        }

        var builder = new System.Text.StringBuilder();
        for (int i = start; i < start + count; i++) {
            if (i > start) {
                builder.Append(separator);
            }
            GetSlice(i, out var text, out int sliceStart, out int sliceEnd);
            builder.Append(text, sliceStart, sliceEnd - sliceStart);
        }
        return builder.ToString();
    }

    IEnumerator<string> IEnumerable<string>.GetEnumerator() {
        for (int i = 0; i < Length; i++) {
            yield return this[i];
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => ((IEnumerable<string>)this).GetEnumerator();
}

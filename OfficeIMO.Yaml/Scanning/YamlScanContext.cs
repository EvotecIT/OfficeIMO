namespace OfficeIMO.Yaml;

/// <summary>
/// Scanner state that persists between <see cref="YamlReader.Read"/> calls. The <see cref="YamlScanner"/> ref struct
/// is rebuilt over the current text span on each call and keeps everything else here.
/// </summary>
internal sealed class YamlScanContext {
    private const int InitialBufferSize = 16 * 1024;

    // Input. In-memory text has no source and is final from the start.
    public YamlCharSource? Source;
    public char[]? Buffer;
    public int Length;
    public int Base;
    public bool IsFinal = true;
    public long MaxInput = long.MaxValue;

    // Position of the next unscanned character, relative to the current text span.
    public int Pos;
    public int Line;
    public int Column;

    // Block and flow structure.
    public int FlowLevel;
    public int Indent = -1;
    public int[] Indents;
    public int IndentCount;
    public YamlSimpleKey[] SimpleKeys;
    public bool SimpleKeyAllowed;
    public bool StreamStartProduced;
    public bool StreamEndProduced;
    public bool InDocument;
    public bool AdjacentValue;
    public bool TabBefore;

    // Token queue.
    public YamlToken[] Tokens;
    public int TokenHead;
    public int TokenCount;
    public int TokensParsed;

    // End of the last token with a source span (absolute offset; zero-based line and column).
    public int LastEnd;
    public int LastEndLine;
    public int LastEndColumn;
    /// <summary>Line on which the last real token ended; tokens on a later line are the first on their line.</summary>
    public int LastLine = -1;

    // Absolute offset of the last stale implicit-key check; the check is repeated only after the position moves.
    public int StaleCheckedAt = -1;

    // The line of the last full stale check and the oldest offset of an implicit-key candidate since then. While the
    // line is unchanged and that candidate is within the key length limit, no candidate can be stale.
    public int StaleLine = -1;
    public int StaleOldest = int.MaxValue;

    // Comments, recorded only when requested.
    public bool IncludeComments;
    public YamlComment[] Comments;
    public int CommentHead;
    public int CommentCount;

    // Leniency. A finding is recorded once per outermost flow collection and once per quoted scalar.
    public bool Strict;
    public bool FlowLenientReported;
    public List<YamlLenientFinding>? Findings;

    // Limits.
    public int MaxDepth;
    public int MaxScalarLength;
    public CancellationToken Cancellation;
    public int CancelCountdown = CancelInterval;
    public const int CancelInterval = 512;

    /// <summary>Absolute offset before which the current <see cref="YamlReader.Read"/> no longer needs the text, or -1.</summary>
    public int RetainMark = -1;

    /// <summary>
    /// The structure arrays of the last released context on this thread. Reading many small documents otherwise spends
    /// a noticeable share of each read renting and returning them; arrays that grew past <see cref="SpareLimit"/> go
    /// back to the pool instead.
    /// </summary>
    [ThreadStatic] private static ScanArrays? t_spare;
    private const int SpareLimit = 64;
    private ScanArrays? _arrays;

    private sealed class ScanArrays {
        public int[] Indents = Array.Empty<int>();
        public YamlSimpleKey[] SimpleKeys = Array.Empty<YamlSimpleKey>();
        public YamlToken[] Tokens = Array.Empty<YamlToken>();
    }

    public YamlScanContext(YamlReaderOptions options, CancellationToken cancellationToken) {
        MaxDepth = options.MaxDepth;
        MaxScalarLength = options.MaxScalarLength;
        IncludeComments = options.IncludeComments;
        Strict = options.Strict;
        Cancellation = cancellationToken;
        ScanArrays? spare = t_spare;
        if (spare is not null) {
            t_spare = null;
            _arrays = spare;
            Indents = spare.Indents;
            SimpleKeys = spare.SimpleKeys;
            Tokens = spare.Tokens;
        } else {
            Indents = ArrayPool<int>.Shared.Rent(16);
            SimpleKeys = ArrayPool<YamlSimpleKey>.Shared.Rent(8);
            Tokens = ArrayPool<YamlToken>.Shared.Rent(16);
        }
        Comments = options.IncludeComments ? ArrayPool<YamlComment>.Shared.Rent(8) : Array.Empty<YamlComment>();
    }

    /// <summary>Starts buffered input from a character source.</summary>
    public void AttachSource(YamlCharSource source, long maxInput) {
        Source = source;
        Buffer = ArrayPool<char>.Shared.Rent(InitialBufferSize);
        Length = 0;
        IsFinal = false;
        MaxInput = maxInput;
    }

    /// <summary>Appends more input, growing the buffer when it is full. Never moves retained text, so positions stay valid.</summary>
    public bool Fill() {
        if (IsFinal) return false;
        char[] buffer = Buffer!;
        if (Length == buffer.Length) {
            // The window only grows while one token (or the lookahead behind it) is longer than the buffer, so a window
            // past twice the scalar limit means a scalar beyond the limit is still being scanned.
            if (buffer.Length >= (long)MaxScalarLength * 2 + InitialBufferSize * 4) {
                throw new YamlLimitExceededException(nameof(YamlReaderOptions.MaxScalarLength), MaxScalarLength, new YamlPosition((long)Base + Pos, Line + 1, Column + 1));
            }
            char[] larger = ArrayPool<char>.Shared.Rent(buffer.Length * 2);
            Array.Copy(buffer, larger, Length);
            ArrayPool<char>.Shared.Return(buffer);
            Buffer = buffer = larger;
        }
        Cancellation.ThrowIfCancellationRequested();
        int read = Source!.Read(buffer, Length, buffer.Length - Length);
        if (read <= 0) {
            IsFinal = true;
            return false;
        }
        Length += read;
        if (Source.Consumed > MaxInput) {
            throw new YamlLimitExceededException(nameof(YamlReaderOptions.MaxInputBytes), MaxInput, new YamlPosition((long)Base + Length, Line + 1, Column + 1));
        }
        if ((long)Base + Length >= int.MaxValue - 1) {
            throw new YamlLimitExceededException("MaxInputLength", int.MaxValue - 1, new YamlPosition((long)Base + Length, Line + 1, Column + 1));
        }
        return true;
    }

    /// <summary>
    /// Drops text before the earliest position still needed, at a point where no token is half scanned.
    /// Only moves text when at least half the buffer can be reclaimed, which keeps copying amortized.
    /// </summary>
    public void Compact() {
        if (Source is null) return;
        int keep = Base + Pos - 1;
        if (TokenCount > 0 && Tokens[TokenHead].Start < keep) keep = Tokens[TokenHead].Start;
        if (CommentCount > 0 && Comments[CommentHead].Start < keep) keep = Comments[CommentHead].Start;
        if (RetainMark >= 0 && RetainMark < keep) keep = RetainMark;
        int delta = keep - Base;
        if (delta <= 0 || delta < Buffer!.Length / 2) return;
        Array.Copy(Buffer, delta, Buffer, 0, Length - delta);
        Length -= delta;
        Base += delta;
        Pos -= delta;
    }

    public void AddFinding(YamlLenientKind kind, int offset, int line, int column) =>
        (Findings ??= new List<YamlLenientFinding>()).Add(new YamlLenientFinding(kind, offset, line, column));

    public void Release() {
        Source?.Release();
        if (Buffer is not null) { ArrayPool<char>.Shared.Return(Buffer); Buffer = null; }
        if (Indents.Length > 0 && Indents.Length <= SpareLimit && SimpleKeys.Length <= SpareLimit && Tokens.Length <= SpareLimit && t_spare is null) {
            // The arrays hold only offsets and flags, so the next reader on this thread takes them as they are.
            ScanArrays spare = _arrays ?? new ScanArrays();
            spare.Indents = Indents;
            spare.SimpleKeys = SimpleKeys;
            spare.Tokens = Tokens;
            t_spare = spare;
            _arrays = null;
            Indents = Array.Empty<int>();
            SimpleKeys = Array.Empty<YamlSimpleKey>();
            Tokens = Array.Empty<YamlToken>();
        }
        if (Indents.Length > 0) { ArrayPool<int>.Shared.Return(Indents); Indents = Array.Empty<int>(); }
        if (SimpleKeys.Length > 0) { ArrayPool<YamlSimpleKey>.Shared.Return(SimpleKeys); SimpleKeys = Array.Empty<YamlSimpleKey>(); }
        if (Tokens.Length > 0) { ArrayPool<YamlToken>.Shared.Return(Tokens); Tokens = Array.Empty<YamlToken>(); }
        if (Comments.Length > 0) { ArrayPool<YamlComment>.Shared.Return(Comments); Comments = Array.Empty<YamlComment>(); }
    }

    public static T[] Grow<T>(T[] array, int used) {
        T[] larger = ArrayPool<T>.Shared.Rent(Math.Max(array.Length * 2, 16));
        Array.Copy(array, larger, used);
        if (array.Length > 0) ArrayPool<T>.Shared.Return(array);
        return larger;
    }
}

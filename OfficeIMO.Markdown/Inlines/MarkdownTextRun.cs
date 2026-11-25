namespace OfficeIMO.Markdown;

/// <summary>
/// Plain text run.
/// </summary>
public sealed class MarkdownTextRun : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, ILiteralTextMarkdownInline, IHtmlAppendableMarkdownInline {
    // The text, or for a run parsed from an escaped character, its escape details carrying the text.
    // Escaped runs are rare, so plain runs keep one field instead of a text and a metadata reference.
    private object _value;

    // A parsed run can point into the inline text it was read from instead of copying it. Then _value
    // holds that text and _slice packs the run's start (upper 15 bits) and length (lower 16 bits); the
    // first read of Text replaces _value with the copied string. _slice is negative for every other run
    // and never changes, and a string of exactly the slice length is always the run's own text (a
    // source of that length starts at 0), so a read racing that replacement returns the same text.
    // One int keeps the run the size it was before slices.
    private readonly int _slice = -1;
    private const int MaxSliceStart = 0x7FFF;
    private const int MaxSliceLength = 0xFFFF;

    // Source details of the parse; reading them captures deferred source information.
    private MarkdownTextRunSyntaxMetadata? SourceMetadata {
        get {
            EnsureSourceInformation();
            return _value as MarkdownTextRunSyntaxMetadata;
        }
        set {
            EnsureSourceInformation();
            SetMetadata(value);
        }
    }

    /// <summary>Text content.</summary>
    public string Text {
        get {
            object value = _value;
            if (value is string text) {
                int slice = _slice;
                return slice < 0 || text.Length == (slice & MaxSliceLength) ? text : MaterializeSlice(text, slice);
            }

            return ((MarkdownTextRunSyntaxMetadata)value).Text;
        }
    }

    private static int SliceStart(int slice) => slice >> 16;

    private static int SliceLength(int slice) => slice & MaxSliceLength;

    private string MaterializeSlice(string source, int slice) {
        string text = source.Substring(SliceStart(slice), SliceLength(slice));
        _value = text;
        return text;
    }

    // The run's text as a range of a string, without copying a slice that was not read yet.
    private string GetTextRange(out int start, out int length) {
        int slice = _slice;
        if (slice >= 0 && _value is string source && source.Length != SliceLength(slice)) {
            start = SliceStart(slice);
            length = SliceLength(slice);
            return source;
        }

        string text = Text;
        start = 0;
        length = text.Length;
        return text;
    }

    private void SetMetadata(MarkdownTextRunSyntaxMetadata? metadata) {
        string text = Text;
        if (metadata == null) {
            _value = text;
            return;
        }

        metadata.Text = text;
        _value = metadata;
    }
    /// <summary>Backslash escape marker text when this run was parsed from an escaped character.</summary>
    public string? EscapeMarker {
        get => SourceMetadata?.EscapeMarker;
        internal set {
            if (value != null) SyntaxMetadata.EscapeMarker = value;
            else if (SourceMetadata != null) SourceMetadata.EscapeMarker = null;
        }
    }
    /// <summary>Source span for the backslash escape marker when this run was parsed from markdown.</summary>
    public MarkdownSourceSpan? EscapeMarkerSourceSpan {
        get => SourceMetadata?.EscapeMarkerSourceSpan;
        internal set {
            if (value.HasValue) SyntaxMetadata.EscapeMarkerSourceSpan = value;
            else if (SourceMetadata != null) SourceMetadata.EscapeMarkerSourceSpan = null;
        }
    }
    /// <summary>Escaped character text when this run was parsed from an escaped character.</summary>
    public string? EscapedCharacter {
        get => SourceMetadata?.EscapedCharacter;
        internal set {
            if (value != null) SyntaxMetadata.EscapedCharacter = value;
            else if (SourceMetadata != null) SourceMetadata.EscapedCharacter = null;
        }
    }
    /// <summary>Source span for the escaped character when this run was parsed from markdown.</summary>
    public MarkdownSourceSpan? EscapedCharacterSourceSpan {
        get => SourceMetadata?.EscapedCharacterSourceSpan;
        internal set {
            if (value.HasValue) SyntaxMetadata.EscapedCharacterSourceSpan = value;
            else if (SourceMetadata != null) SourceMetadata.EscapedCharacterSourceSpan = null;
        }
    }
    /// <summary>Creates a plain text run.</summary>
    public MarkdownTextRun(string text) { _value = text ?? string.Empty; }

    // A run of the length characters of source at start, copied only when Text is read (or now, for a
    // slice too far into a long text to pack).
    internal MarkdownTextRun(string source, int start, int length) {
        if (start == 0 && length == source.Length) {
            _value = source;
        } else if (start <= MaxSliceStart && length <= MaxSliceLength) {
            _value = source;
            _slice = (start << 16) | length;
        } else {
            _value = source.Substring(start, length);
        }
    }

    internal void SetMarkdownSyntaxMetadataSpans(
        string? escapeMarker,
        MarkdownSourceSpan? escapeMarkerSourceSpan,
        string? escapedCharacter,
        MarkdownSourceSpan? escapedCharacterSourceSpan) {
        if (escapeMarker == null
            && !escapeMarkerSourceSpan.HasValue
            && escapedCharacter == null
            && !escapedCharacterSourceSpan.HasValue) {
            SourceMetadata = null;
            return;
        }

        SourceMetadata = new MarkdownTextRunSyntaxMetadata {
            EscapeMarker = escapeMarker,
            EscapeMarkerSourceSpan = escapeMarkerSourceSpan,
            EscapedCharacter = escapedCharacter,
            EscapedCharacterSourceSpan = escapedCharacterSourceSpan
        };
    }

    internal string RenderMarkdown() => MarkdownEscaper.EscapeTextForWriter(Text);
    internal string RenderHtml() => HtmlTextEncoder.Encode(Text, HtmlRenderContext.Options);
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IHtmlAppendableMarkdownInline.AppendHtml(StringBuilder builder) {
        string text = GetTextRange(out int start, out int length);
        HtmlTextEncoder.Append(builder, text, start, length, HtmlRenderContext.Options);
    }
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) {
        string text = GetTextRange(out int start, out int length);
        sb.Append(text, start, length);
    }

    internal override void AdoptCapturedSourceInformation(MarkdownObject captured, MarkdownCapturedObjectMap map) {
        // Copy the details rather than share them: the holder also carries this run's text.
        SetMetadata(((MarkdownTextRun)captured)._value is MarkdownTextRunSyntaxMetadata source
            ? new MarkdownTextRunSyntaxMetadata {
                EscapeMarker = source.EscapeMarker,
                EscapeMarkerSourceSpan = source.EscapeMarkerSourceSpan,
                EscapedCharacter = source.EscapedCharacter,
                EscapedCharacterSourceSpan = source.EscapedCharacterSourceSpan
            }
            : null);
    }

    private MarkdownTextRunSyntaxMetadata SyntaxMetadata =>
        SourceMetadata ??= new MarkdownTextRunSyntaxMetadata();
}

internal sealed class MarkdownTextRunSyntaxMetadata {
    internal string Text = string.Empty;
    internal string? EscapeMarker;
    internal MarkdownSourceSpan? EscapeMarkerSourceSpan;
    internal string? EscapedCharacter;
    internal MarkdownSourceSpan? EscapedCharacterSourceSpan;
}

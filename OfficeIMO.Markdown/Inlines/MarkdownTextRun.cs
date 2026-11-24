namespace OfficeIMO.Markdown;

/// <summary>
/// Plain text run.
/// </summary>
public sealed class MarkdownTextRun : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, ILiteralTextMarkdownInline, IHtmlAppendableMarkdownInline {
    // The text, or for a run parsed from an escaped character, its escape details carrying the text.
    // Escaped runs are rare, so plain runs keep one field instead of a text and a metadata reference.
    private object _value;

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
            return value as string ?? ((MarkdownTextRunSyntaxMetadata)value).Text;
        }
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
    void IHtmlAppendableMarkdownInline.AppendHtml(StringBuilder builder) => HtmlTextEncoder.Append(builder, Text, HtmlRenderContext.Options);
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);

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

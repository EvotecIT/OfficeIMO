namespace OfficeIMO.Markdown;

/// <summary>
/// Plain text run.
/// </summary>
public sealed class MarkdownTextRun : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, ILiteralTextMarkdownInline, IHtmlAppendableMarkdownInline {
    private MarkdownTextRunSyntaxMetadata? _syntaxMetadata;

    // Source details of the parse; reading them captures deferred source information.
    private MarkdownTextRunSyntaxMetadata? SourceMetadata {
        get => ReadSourceInformation(ref _syntaxMetadata);
        set => WriteSourceInformation(ref _syntaxMetadata, value);
    }

    /// <summary>Text content.</summary>
    public string Text { get; }
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
    public MarkdownTextRun(string text) { Text = text ?? string.Empty; }

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
        _syntaxMetadata = ((MarkdownTextRun)captured)._syntaxMetadata;
    }

    private MarkdownTextRunSyntaxMetadata SyntaxMetadata =>
        SourceMetadata ??= new MarkdownTextRunSyntaxMetadata();
}

internal sealed class MarkdownTextRunSyntaxMetadata {
    internal string? EscapeMarker;
    internal MarkdownSourceSpan? EscapeMarkerSourceSpan;
    internal string? EscapedCharacter;
    internal MarkdownSourceSpan? EscapedCharacterSourceSpan;
}

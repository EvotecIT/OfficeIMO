namespace OfficeIMO.Markdown;

/// <summary>
/// Plain text run.
/// </summary>
public sealed class MarkdownTextRun : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, ILiteralTextMarkdownInline {
    private MarkdownTextRunSyntaxMetadata? _syntaxMetadata;

    /// <summary>Text content.</summary>
    public string Text { get; }
    /// <summary>Backslash escape marker text when this run was parsed from an escaped character.</summary>
    public string? EscapeMarker {
        get => _syntaxMetadata?.EscapeMarker;
        internal set {
            if (value != null) SyntaxMetadata.EscapeMarker = value;
            else if (_syntaxMetadata != null) _syntaxMetadata.EscapeMarker = null;
        }
    }
    /// <summary>Source span for the backslash escape marker when this run was parsed from markdown.</summary>
    public MarkdownSourceSpan? EscapeMarkerSourceSpan {
        get => _syntaxMetadata?.EscapeMarkerSourceSpan;
        internal set {
            if (value.HasValue) SyntaxMetadata.EscapeMarkerSourceSpan = value;
            else if (_syntaxMetadata != null) _syntaxMetadata.EscapeMarkerSourceSpan = null;
        }
    }
    /// <summary>Escaped character text when this run was parsed from an escaped character.</summary>
    public string? EscapedCharacter {
        get => _syntaxMetadata?.EscapedCharacter;
        internal set {
            if (value != null) SyntaxMetadata.EscapedCharacter = value;
            else if (_syntaxMetadata != null) _syntaxMetadata.EscapedCharacter = null;
        }
    }
    /// <summary>Source span for the escaped character when this run was parsed from markdown.</summary>
    public MarkdownSourceSpan? EscapedCharacterSourceSpan {
        get => _syntaxMetadata?.EscapedCharacterSourceSpan;
        internal set {
            if (value.HasValue) SyntaxMetadata.EscapedCharacterSourceSpan = value;
            else if (_syntaxMetadata != null) _syntaxMetadata.EscapedCharacterSourceSpan = null;
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
            _syntaxMetadata = null;
            return;
        }

        _syntaxMetadata = new MarkdownTextRunSyntaxMetadata {
            EscapeMarker = escapeMarker,
            EscapeMarkerSourceSpan = escapeMarkerSourceSpan,
            EscapedCharacter = escapedCharacter,
            EscapedCharacterSourceSpan = escapedCharacterSourceSpan
        };
    }

    internal string RenderMarkdown() => MarkdownEscaper.EscapeText(Text);
    internal string RenderHtml() => HtmlTextEncoder.Encode(Text, HtmlRenderContext.Options);
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);

    private MarkdownTextRunSyntaxMetadata SyntaxMetadata =>
        _syntaxMetadata ??= new MarkdownTextRunSyntaxMetadata();
}

internal sealed class MarkdownTextRunSyntaxMetadata {
    internal string? EscapeMarker;
    internal MarkdownSourceSpan? EscapeMarkerSourceSpan;
    internal string? EscapedCharacter;
    internal MarkdownSourceSpan? EscapedCharacterSourceSpan;
}

using System.Runtime.CompilerServices;

namespace OfficeIMO.Markdown;

internal static class MarkdownInlineMetadataSourceSpans {
    private sealed class LinkState {
        public MarkdownSourceSpan? TargetSpan;
        public MarkdownSourceSpan? TitleSpan;
        public MarkdownSourceSpan? HtmlTargetSpan;
        public MarkdownSourceSpan? HtmlRelSpan;
        public string? AutolinkLiteral;
    }

    private sealed class LinkHolder {
        public LinkState? State;
    }

    private class ImageState {
        public MarkdownSourceSpan? AltSpan;
        public MarkdownSourceSpan? SourceSpan;
        public MarkdownSourceSpan? TitleSpan;
    }

    private class ImageHolder {
        public ImageState? State;
    }

    private sealed class ImageLinkState : ImageState {
        public MarkdownSourceSpan? LinkTargetSpan;
        public MarkdownSourceSpan? LinkTitleSpan;
    }

    private sealed class ImageLinkHolder : ImageHolder {
        public new ImageLinkState? State;
    }

    private sealed class FormattingMarkerState {
        public string OpeningMarker = string.Empty;
        public MarkdownSourceSpan? OpeningMarkerSpan;
        public string SeparatorMarker = string.Empty;
        public MarkdownSourceSpan? SeparatorMarkerSpan;
        public string ClosingMarker = string.Empty;
        public MarkdownSourceSpan? ClosingMarkerSpan;
    }

    private sealed class FormattingMarkerHolder {
        public FormattingMarkerState? State;
    }

    private sealed class CodeSpanState {
        public MarkdownSourceSpan? ContentSpan;
    }

    private sealed class CodeSpanHolder {
        public CodeSpanState? State;
    }

    private sealed class EscapedTextState {
        public string EscapeMarker = string.Empty;
        public MarkdownSourceSpan? EscapeMarkerSpan;
        public string EscapedCharacter = string.Empty;
        public MarkdownSourceSpan? EscapedCharacterSpan;
    }

    private sealed class EscapedTextHolder {
        public EscapedTextState? State;
    }

    private sealed class DecodedEntityState {
        public string SourceText = string.Empty;
        public MarkdownSourceSpan? SourceTextSpan;
    }

    private sealed class DecodedEntityHolder {
        public DecodedEntityState? State;
    }

    private sealed class HardBreakMarkerState {
        public string Marker = string.Empty;
        public MarkdownSourceSpan? MarkerSpan;
    }

    private sealed class HardBreakMarkerHolder {
        public HardBreakMarkerState? State;
    }

    private sealed class AbbreviationState {
        public MarkdownSourceSpan? TextSpan;
        public MarkdownSourceSpan? TitleSpan;
    }

    private sealed class AbbreviationHolder {
        public AbbreviationState? State;
    }

    // These tables hold weak references to markdown inline keys, so entries disappear when the
    // owning inline objects are no longer referenced by the parse result or callers.
    private static readonly ConditionalWeakTable<LinkInline, LinkHolder> _linkSpans = new();
    private static readonly ConditionalWeakTable<ImageInline, ImageHolder> _imageSpans = new();
    private static readonly ConditionalWeakTable<ImageLinkInline, ImageLinkHolder> _imageLinkSpans = new();
    private static readonly ConditionalWeakTable<MarkdownInline, FormattingMarkerHolder> _formattingMarkerSpans = new();
    private static readonly ConditionalWeakTable<CodeSpanInline, CodeSpanHolder> _codeSpanSpans = new();
    private static readonly ConditionalWeakTable<TextRun, EscapedTextHolder> _escapedTextSpans = new();
    private static readonly ConditionalWeakTable<DecodedHtmlEntityTextRun, DecodedEntityHolder> _decodedEntitySpans = new();
    private static readonly ConditionalWeakTable<HardBreakInline, HardBreakMarkerHolder> _hardBreakMarkerSpans = new();
    private static readonly ConditionalWeakTable<AbbreviationInline, AbbreviationHolder> _abbreviationSpans = new();

    internal static void SetLinkParts(
        LinkInline? inline,
        MarkdownSourceSpan? targetSpan,
        MarkdownSourceSpan? titleSpan,
        MarkdownSourceSpan? htmlTargetSpan = null,
        MarkdownSourceSpan? htmlRelSpan = null,
        string? autolinkLiteral = null) {
        if (inline == null) {
            return;
        }

        if (!targetSpan.HasValue && !titleSpan.HasValue && !htmlTargetSpan.HasValue && !htmlRelSpan.HasValue && string.IsNullOrEmpty(autolinkLiteral)) {
            return;
        }

        var holder = _linkSpans.GetValue(inline, static _ => new LinkHolder());
        holder.State = new LinkState {
            TargetSpan = targetSpan,
            TitleSpan = titleSpan,
            HtmlTargetSpan = htmlTargetSpan,
            HtmlRelSpan = htmlRelSpan,
            AutolinkLiteral = autolinkLiteral
        };
        inline.SetMarkdownSyntaxMetadataSpans(targetSpan, titleSpan, htmlTargetSpan, htmlRelSpan);
    }

    internal static MarkdownSourceSpan? GetLinkTargetSpan(LinkInline? inline) =>
        inline?.UrlSourceSpan
        ?? (inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.State?.TargetSpan : null);

    internal static MarkdownSourceSpan? GetLinkTitleSpan(LinkInline? inline) =>
        inline?.TitleSourceSpan
        ?? (inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.State?.TitleSpan : null);

    internal static MarkdownSourceSpan? GetLinkHtmlTargetSpan(LinkInline? inline) =>
        inline?.HtmlTargetSourceSpan
        ?? (inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.State?.HtmlTargetSpan : null);

    internal static MarkdownSourceSpan? GetLinkHtmlRelSpan(LinkInline? inline) =>
        inline?.HtmlRelSourceSpan
        ?? (inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.State?.HtmlRelSpan : null);

    internal static string? GetAutolinkLiteral(LinkInline? inline) =>
        inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.State?.AutolinkLiteral : null;

    internal static void SetImageParts(
        ImageInline? inline,
        MarkdownSourceSpan? altSpan,
        MarkdownSourceSpan? sourceSpan,
        MarkdownSourceSpan? titleSpan = null) {
        if (inline == null) {
            return;
        }

        if (!altSpan.HasValue && !sourceSpan.HasValue && !titleSpan.HasValue) {
            return;
        }

        var holder = _imageSpans.GetValue(inline, static _ => new ImageHolder());
        holder.State = new ImageState {
            AltSpan = altSpan,
            SourceSpan = sourceSpan,
            TitleSpan = titleSpan
        };
        inline.SetMarkdownSyntaxMetadataSpans(altSpan, sourceSpan, titleSpan);
    }

    internal static void SetImageLinkParts(
        ImageLinkInline? inline,
        MarkdownSourceSpan? altSpan,
        MarkdownSourceSpan? sourceSpan,
        MarkdownSourceSpan? imageTitleSpan,
        MarkdownSourceSpan? linkTargetSpan,
        MarkdownSourceSpan? linkTitleSpan = null) {
        if (inline == null) {
            return;
        }

        if (!altSpan.HasValue &&
            !sourceSpan.HasValue &&
            !imageTitleSpan.HasValue &&
            !linkTargetSpan.HasValue &&
            !linkTitleSpan.HasValue) {
            return;
        }

        var holder = _imageLinkSpans.GetValue(inline, static _ => new ImageLinkHolder());
        holder.State = new ImageLinkState {
            AltSpan = altSpan,
            SourceSpan = sourceSpan,
            TitleSpan = imageTitleSpan,
            LinkTargetSpan = linkTargetSpan,
            LinkTitleSpan = linkTitleSpan
        };
        inline.SetMarkdownSyntaxMetadataSpans(altSpan, sourceSpan, imageTitleSpan, linkTargetSpan, linkTitleSpan);
    }

    internal static MarkdownSourceSpan? GetImageAltSpan(ImageInline? inline) =>
        inline?.AltSourceSpan
        ?? (inline != null && _imageSpans.TryGetValue(inline, out var holder) ? holder.State?.AltSpan : null);

    internal static MarkdownSourceSpan? GetImageSourceSpan(ImageInline? inline) =>
        inline?.SrcSourceSpan
        ?? (inline != null && _imageSpans.TryGetValue(inline, out var holder) ? holder.State?.SourceSpan : null);

    internal static MarkdownSourceSpan? GetImageTitleSpan(ImageInline? inline) =>
        inline?.TitleSourceSpan
        ?? (inline != null && _imageSpans.TryGetValue(inline, out var holder) ? holder.State?.TitleSpan : null);

    internal static MarkdownSourceSpan? GetImageAltSpan(ImageLinkInline? inline) =>
        inline?.AltSourceSpan
        ?? (inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.State?.AltSpan : null);

    internal static MarkdownSourceSpan? GetImageSourceSpan(ImageLinkInline? inline) =>
        inline?.ImageUrlSourceSpan
        ?? (inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.State?.SourceSpan : null);

    internal static MarkdownSourceSpan? GetImageTitleSpan(ImageLinkInline? inline) =>
        inline?.TitleSourceSpan
        ?? (inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.State?.TitleSpan : null);

    internal static MarkdownSourceSpan? GetImageLinkTargetSpan(ImageLinkInline? inline) =>
        inline?.LinkUrlSourceSpan
        ?? (inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.State?.LinkTargetSpan : null);

    internal static MarkdownSourceSpan? GetImageLinkTitleSpan(ImageLinkInline? inline) =>
        inline?.LinkTitleSourceSpan
        ?? (inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.State?.LinkTitleSpan : null);

    internal static void SetFormattingMarkers(
        MarkdownInline? inline,
        string openingMarker,
        MarkdownSourceSpan? openingMarkerSpan,
        string closingMarker,
        MarkdownSourceSpan? closingMarkerSpan,
        string? separatorMarker = null,
        MarkdownSourceSpan? separatorMarkerSpan = null) {
        if (inline == null) {
            return;
        }

        if (string.IsNullOrEmpty(openingMarker) &&
            string.IsNullOrEmpty(separatorMarker) &&
            string.IsNullOrEmpty(closingMarker) &&
            !openingMarkerSpan.HasValue &&
            !separatorMarkerSpan.HasValue &&
            !closingMarkerSpan.HasValue) {
            return;
        }

        var holder = _formattingMarkerSpans.GetValue(inline, static _ => new FormattingMarkerHolder());
        holder.State = new FormattingMarkerState {
            OpeningMarker = openingMarker ?? string.Empty,
            OpeningMarkerSpan = openingMarkerSpan,
            SeparatorMarker = separatorMarker ?? string.Empty,
            SeparatorMarkerSpan = separatorMarkerSpan,
            ClosingMarker = closingMarker ?? string.Empty,
            ClosingMarkerSpan = closingMarkerSpan
        };
    }

    internal static string? GetOpeningMarker(MarkdownInline? inline) =>
        inline != null && _formattingMarkerSpans.TryGetValue(inline, out var holder) ? holder.State?.OpeningMarker : null;

    internal static MarkdownSourceSpan? GetOpeningMarkerSpan(MarkdownInline? inline) =>
        inline != null && _formattingMarkerSpans.TryGetValue(inline, out var holder) ? holder.State?.OpeningMarkerSpan : null;

    internal static string? GetSeparatorMarker(MarkdownInline? inline) =>
        inline != null && _formattingMarkerSpans.TryGetValue(inline, out var holder) ? holder.State?.SeparatorMarker : null;

    internal static MarkdownSourceSpan? GetSeparatorMarkerSpan(MarkdownInline? inline) =>
        inline != null && _formattingMarkerSpans.TryGetValue(inline, out var holder) ? holder.State?.SeparatorMarkerSpan : null;

    internal static string? GetClosingMarker(MarkdownInline? inline) =>
        inline != null && _formattingMarkerSpans.TryGetValue(inline, out var holder) ? holder.State?.ClosingMarker : null;

    internal static MarkdownSourceSpan? GetClosingMarkerSpan(MarkdownInline? inline) =>
        inline != null && _formattingMarkerSpans.TryGetValue(inline, out var holder) ? holder.State?.ClosingMarkerSpan : null;

    internal static void SetCodeSpanContent(
        CodeSpanInline? inline,
        MarkdownSourceSpan? contentSpan) {
        if (inline == null || !contentSpan.HasValue) {
            return;
        }

        var holder = _codeSpanSpans.GetValue(inline, static _ => new CodeSpanHolder());
        holder.State = new CodeSpanState {
            ContentSpan = contentSpan
        };
    }

    internal static MarkdownSourceSpan? GetCodeSpanContentSpan(CodeSpanInline? inline) =>
        inline != null && _codeSpanSpans.TryGetValue(inline, out var holder) ? holder.State?.ContentSpan : null;

    internal static void SetEscapedText(
        TextRun? inline,
        string escapeMarker,
        MarkdownSourceSpan? escapeMarkerSpan,
        string escapedCharacter,
        MarkdownSourceSpan? escapedCharacterSpan) {
        if (inline == null) {
            return;
        }

        if (string.IsNullOrEmpty(escapeMarker) &&
            string.IsNullOrEmpty(escapedCharacter) &&
            !escapeMarkerSpan.HasValue &&
            !escapedCharacterSpan.HasValue) {
            return;
        }

        var holder = _escapedTextSpans.GetValue(inline, static _ => new EscapedTextHolder());
        holder.State = new EscapedTextState {
            EscapeMarker = escapeMarker ?? string.Empty,
            EscapeMarkerSpan = escapeMarkerSpan,
            EscapedCharacter = escapedCharacter ?? string.Empty,
            EscapedCharacterSpan = escapedCharacterSpan
        };
    }

    internal static string? GetEscapeMarker(TextRun? inline) =>
        inline != null && _escapedTextSpans.TryGetValue(inline, out var holder) ? holder.State?.EscapeMarker : null;

    internal static MarkdownSourceSpan? GetEscapeMarkerSpan(TextRun? inline) =>
        inline != null && _escapedTextSpans.TryGetValue(inline, out var holder) ? holder.State?.EscapeMarkerSpan : null;

    internal static string? GetEscapedCharacter(TextRun? inline) =>
        inline != null && _escapedTextSpans.TryGetValue(inline, out var holder) ? holder.State?.EscapedCharacter : null;

    internal static MarkdownSourceSpan? GetEscapedCharacterSpan(TextRun? inline) =>
        inline != null && _escapedTextSpans.TryGetValue(inline, out var holder) ? holder.State?.EscapedCharacterSpan : null;

    internal static void SetDecodedEntity(
        DecodedHtmlEntityTextRun? inline,
        string sourceText,
        MarkdownSourceSpan? sourceTextSpan) {
        if (inline == null) {
            return;
        }

        if (string.IsNullOrEmpty(sourceText) && !sourceTextSpan.HasValue) {
            return;
        }

        var holder = _decodedEntitySpans.GetValue(inline, static _ => new DecodedEntityHolder());
        holder.State = new DecodedEntityState {
            SourceText = sourceText ?? string.Empty,
            SourceTextSpan = sourceTextSpan
        };
    }

    internal static string? GetDecodedEntitySourceText(DecodedHtmlEntityTextRun? inline) =>
        inline != null && _decodedEntitySpans.TryGetValue(inline, out var holder) ? holder.State?.SourceText : null;

    internal static MarkdownSourceSpan? GetDecodedEntitySourceTextSpan(DecodedHtmlEntityTextRun? inline) =>
        inline != null && _decodedEntitySpans.TryGetValue(inline, out var holder) ? holder.State?.SourceTextSpan : null;

    internal static void SetHardBreakMarker(
        HardBreakInline? inline,
        string marker,
        MarkdownSourceSpan? markerSpan) {
        if (inline == null) {
            return;
        }

        if (string.IsNullOrEmpty(marker) && !markerSpan.HasValue) {
            return;
        }

        var holder = _hardBreakMarkerSpans.GetValue(inline, static _ => new HardBreakMarkerHolder());
        holder.State = new HardBreakMarkerState {
            Marker = marker ?? string.Empty,
            MarkerSpan = markerSpan
        };
    }

    internal static string? GetHardBreakMarker(HardBreakInline? inline) =>
        inline != null && _hardBreakMarkerSpans.TryGetValue(inline, out var holder) ? holder.State?.Marker : null;

    internal static MarkdownSourceSpan? GetHardBreakMarkerSpan(HardBreakInline? inline) =>
        inline != null && _hardBreakMarkerSpans.TryGetValue(inline, out var holder) ? holder.State?.MarkerSpan : null;

    internal static void SetAbbreviationParts(
        AbbreviationInline? inline,
        MarkdownSourceSpan? textSpan,
        MarkdownSourceSpan? titleSpan) {
        if (inline == null || (!textSpan.HasValue && !titleSpan.HasValue)) {
            return;
        }

        var holder = _abbreviationSpans.GetValue(inline, static _ => new AbbreviationHolder());
        holder.State = new AbbreviationState {
            TextSpan = textSpan,
            TitleSpan = titleSpan
        };
    }

    internal static MarkdownSourceSpan? GetAbbreviationTextSpan(AbbreviationInline? inline) =>
        inline != null && _abbreviationSpans.TryGetValue(inline, out var holder) ? holder.State?.TextSpan : null;

    internal static MarkdownSourceSpan? GetAbbreviationTitleSpan(AbbreviationInline? inline) =>
        inline != null && _abbreviationSpans.TryGetValue(inline, out var holder) ? holder.State?.TitleSpan : null;
}

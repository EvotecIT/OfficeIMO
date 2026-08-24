using System.Runtime.CompilerServices;

namespace OfficeIMO.Markdown;

internal static class MarkdownInlineMetadataSourceSpans {
    private static readonly ConditionalWeakTable<MarkdownInline, MarkdownInlineAuxiliarySyntaxMetadata> _auxiliaryMetadata = new();

    private static MarkdownInlineAuxiliarySyntaxMetadata GetOrCreateAuxiliaryMetadata(MarkdownInline inline) {
        if (inline is IMarkdownInlineAuxiliarySyntaxMetadataOwner owner) {
            return owner.AuxiliarySyntaxMetadata ??= new MarkdownInlineAuxiliarySyntaxMetadata();
        }

        return _auxiliaryMetadata.GetValue(inline, static _ => new MarkdownInlineAuxiliarySyntaxMetadata());
    }

    private static MarkdownInlineAuxiliarySyntaxMetadata? GetAuxiliaryMetadata(MarkdownInline? inline) {
        if (inline is IMarkdownInlineAuxiliarySyntaxMetadataOwner owner) {
            return owner.AuxiliarySyntaxMetadata;
        }

        return inline != null && _auxiliaryMetadata.TryGetValue(inline, out var metadata) ? metadata : null;
    }

    internal static void ReleaseRedundantFormattingMetadata(MarkdownInline inline) {
        if (inline.BoundSyntaxNode == null) {
            return;
        }

        MarkdownInlineAuxiliarySyntaxMetadata? metadata = GetAuxiliaryMetadata(inline);
        if (metadata == null) {
            return;
        }

        if (!string.IsNullOrEmpty(metadata.AutolinkLiteral)) {
            metadata.OpeningMarker = string.Empty;
            metadata.OpeningMarkerSpan = null;
            metadata.SeparatorMarker = string.Empty;
            metadata.SeparatorMarkerSpan = null;
            metadata.ClosingMarker = string.Empty;
            metadata.ClosingMarkerSpan = null;
            return;
        }

        if (inline is IMarkdownInlineAuxiliarySyntaxMetadataOwner owner) {
            owner.AuxiliarySyntaxMetadata = null;
        } else {
            _auxiliaryMetadata.Remove(inline);
        }
    }

    private static MarkdownSyntaxNode? GetBoundChild(MarkdownInline? inline, MarkdownSyntaxKind kind) {
        MarkdownSyntaxNode? syntaxNode = inline?.BoundSyntaxNode;
        if (syntaxNode == null) {
            return null;
        }

        for (int i = 0; i < syntaxNode.Children.Count; i++) {
            if (syntaxNode.Children[i].Kind == kind) {
                return syntaxNode.Children[i];
            }
        }

        return null;
    }

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

        if (!string.IsNullOrEmpty(autolinkLiteral)) {
            GetOrCreateAuxiliaryMetadata(inline).AutolinkLiteral = autolinkLiteral;
        }
        inline.SetMarkdownSyntaxMetadataSpans(targetSpan, titleSpan, htmlTargetSpan, htmlRelSpan);
    }

    internal static MarkdownSourceSpan? GetLinkTargetSpan(LinkInline? inline) =>
        inline?.UrlSourceSpan;

    internal static MarkdownSourceSpan? GetLinkTitleSpan(LinkInline? inline) =>
        inline?.TitleSourceSpan;

    internal static MarkdownSourceSpan? GetLinkHtmlTargetSpan(LinkInline? inline) =>
        inline?.HtmlTargetSourceSpan;

    internal static MarkdownSourceSpan? GetLinkHtmlRelSpan(LinkInline? inline) =>
        inline?.HtmlRelSourceSpan;

    internal static string? GetAutolinkLiteral(LinkInline? inline) =>
        GetAuxiliaryMetadata(inline)?.AutolinkLiteral;

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

        inline.SetMarkdownSyntaxMetadataSpans(altSpan, sourceSpan, imageTitleSpan, linkTargetSpan, linkTitleSpan);
    }

    internal static MarkdownSourceSpan? GetImageAltSpan(ImageInline? inline) =>
        inline?.AltSourceSpan;

    internal static MarkdownSourceSpan? GetImageSourceSpan(ImageInline? inline) =>
        inline?.SrcSourceSpan;

    internal static MarkdownSourceSpan? GetImageTitleSpan(ImageInline? inline) =>
        inline?.TitleSourceSpan;

    internal static MarkdownSourceSpan? GetImageAltSpan(ImageLinkInline? inline) =>
        inline?.AltSourceSpan;

    internal static MarkdownSourceSpan? GetImageSourceSpan(ImageLinkInline? inline) =>
        inline?.ImageUrlSourceSpan;

    internal static MarkdownSourceSpan? GetImageTitleSpan(ImageLinkInline? inline) =>
        inline?.TitleSourceSpan;

    internal static MarkdownSourceSpan? GetImageLinkTargetSpan(ImageLinkInline? inline) =>
        inline?.LinkUrlSourceSpan;

    internal static MarkdownSourceSpan? GetImageLinkTitleSpan(ImageLinkInline? inline) =>
        inline?.LinkTitleSourceSpan;

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

        var metadata = GetOrCreateAuxiliaryMetadata(inline);
        metadata.OpeningMarker = openingMarker ?? string.Empty;
        metadata.OpeningMarkerSpan = openingMarkerSpan;
        metadata.SeparatorMarker = separatorMarker ?? string.Empty;
        metadata.SeparatorMarkerSpan = separatorMarkerSpan;
        metadata.ClosingMarker = closingMarker ?? string.Empty;
        metadata.ClosingMarkerSpan = closingMarkerSpan;
    }

    internal static string? GetOpeningMarker(MarkdownInline? inline) =>
        GetAuxiliaryMetadata(inline)?.OpeningMarker is { Length: > 0 } marker
            ? marker
            : GetBoundChild(inline, MarkdownSyntaxKind.InlineOpeningMarker)?.Literal;

    internal static MarkdownSourceSpan? GetOpeningMarkerSpan(MarkdownInline? inline) =>
        GetAuxiliaryMetadata(inline)?.OpeningMarkerSpan
            ?? GetBoundChild(inline, MarkdownSyntaxKind.InlineOpeningMarker)?.SourceSpan;

    internal static string? GetSeparatorMarker(MarkdownInline? inline) =>
        GetAuxiliaryMetadata(inline)?.SeparatorMarker is { Length: > 0 } marker
            ? marker
            : GetBoundChild(inline, MarkdownSyntaxKind.InlineSeparatorMarker)?.Literal;

    internal static MarkdownSourceSpan? GetSeparatorMarkerSpan(MarkdownInline? inline) =>
        GetAuxiliaryMetadata(inline)?.SeparatorMarkerSpan
            ?? GetBoundChild(inline, MarkdownSyntaxKind.InlineSeparatorMarker)?.SourceSpan;

    internal static string? GetClosingMarker(MarkdownInline? inline) =>
        GetAuxiliaryMetadata(inline)?.ClosingMarker is { Length: > 0 } marker
            ? marker
            : GetBoundChild(inline, MarkdownSyntaxKind.InlineClosingMarker)?.Literal;

    internal static MarkdownSourceSpan? GetClosingMarkerSpan(MarkdownInline? inline) =>
        GetAuxiliaryMetadata(inline)?.ClosingMarkerSpan
            ?? GetBoundChild(inline, MarkdownSyntaxKind.InlineClosingMarker)?.SourceSpan;

    internal static void SetCodeSpanContent(
        CodeSpanInline? inline,
        MarkdownSourceSpan? contentSpan) {
        if (inline == null || !contentSpan.HasValue) {
            return;
        }

        inline.SetMarkdownSyntaxMetadataSpans(contentSpan);
    }

    internal static MarkdownSourceSpan? GetCodeSpanContentSpan(CodeSpanInline? inline) =>
        inline?.ContentSourceSpan;

    internal static void SetEscapedText(
        MarkdownTextRun? inline,
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

        inline.SetMarkdownSyntaxMetadataSpans(escapeMarker, escapeMarkerSpan, escapedCharacter, escapedCharacterSpan);
    }

    internal static string? GetEscapeMarker(MarkdownTextRun? inline) =>
        inline?.EscapeMarker;

    internal static MarkdownSourceSpan? GetEscapeMarkerSpan(MarkdownTextRun? inline) =>
        inline?.EscapeMarkerSourceSpan;

    internal static string? GetEscapedCharacter(MarkdownTextRun? inline) =>
        inline?.EscapedCharacter;

    internal static MarkdownSourceSpan? GetEscapedCharacterSpan(MarkdownTextRun? inline) =>
        inline?.EscapedCharacterSourceSpan;

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

        inline.SetMarkdownSyntaxMetadataSpans(sourceText, sourceTextSpan);
    }

    internal static string? GetDecodedEntitySourceText(DecodedHtmlEntityTextRun? inline) =>
        inline?.SourceText;

    internal static MarkdownSourceSpan? GetDecodedEntitySourceTextSpan(DecodedHtmlEntityTextRun? inline) =>
        inline?.SourceTextSourceSpan;

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

        inline.SetMarkdownSyntaxMetadataSpans(marker, markerSpan);
    }

    internal static string? GetHardBreakMarker(HardBreakInline? inline) =>
        inline?.Marker;

    internal static MarkdownSourceSpan? GetHardBreakMarkerSpan(HardBreakInline? inline) =>
        inline?.MarkerSourceSpan;

    internal static void SetAbbreviationParts(
        AbbreviationInline? inline,
        MarkdownSourceSpan? textSpan,
        MarkdownSourceSpan? titleSpan) {
        if (inline == null || (!textSpan.HasValue && !titleSpan.HasValue)) {
            return;
        }

        inline.SetMarkdownSyntaxMetadataSpans(textSpan, titleSpan);
    }

    internal static MarkdownSourceSpan? GetAbbreviationTextSpan(AbbreviationInline? inline) =>
        inline?.TextSourceSpan;

    internal static MarkdownSourceSpan? GetAbbreviationTitleSpan(AbbreviationInline? inline) =>
        inline?.TitleSourceSpan;
}

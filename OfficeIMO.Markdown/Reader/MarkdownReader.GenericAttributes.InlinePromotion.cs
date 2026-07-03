namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static void PromoteNestedInlineGenericAttributesToParagraph(ParagraphBlock paragraph, MarkdownReaderOptions options) {
        if (paragraph == null ||
            options?.GenericAttributes != true ||
            !paragraph.Attributes.IsEmpty) {
            return;
        }

        if (!TryConsumeNestedInlineGenericAttributes(
            paragraph.Inlines,
            out var attributes,
            out var attributeSourceText,
            out var attributeSpan)) {
            return;
        }

        paragraph.SetAttributes(attributes);
        MarkdownGenericAttributeSourceSpans.Set(paragraph, attributeSourceText, attributeSpan);
    }

    private static bool TryConsumeNestedInlineGenericAttributes(
        InlineSequence sequence,
        out MarkdownAttributeSet attributes,
        out string attributeSourceText,
        out MarkdownSourceSpan? attributeSpan) {
        attributes = MarkdownAttributeSet.Empty;
        attributeSourceText = string.Empty;
        attributeSpan = null;

        if (sequence == null || sequence.Nodes.Count == 0) {
            return false;
        }

        var rewritten = new List<IMarkdownInline>(sequence.Nodes.Count);
        var consumed = false;
        for (var i = 0; i < sequence.Nodes.Count; i++) {
            var node = sequence.Nodes[i];

            if (!consumed &&
                TryConsumeNestedInlineGenericAttributes(node, out var replacement, out attributes, out attributeSourceText, out attributeSpan)) {
                rewritten.Add(replacement ?? node);
                consumed = true;
                continue;
            }

            rewritten.Add(node);
        }

        if (!consumed) {
            return false;
        }

        sequence.ReplaceItems(rewritten);
        return true;
    }

    private static bool TryConsumeNestedInlineGenericAttributes(
        IMarkdownInline inline,
        out IMarkdownInline? replacement,
        out MarkdownAttributeSet attributes,
        out string attributeSourceText,
        out MarkdownSourceSpan? attributeSpan) {
        replacement = null;
        attributes = MarkdownAttributeSet.Empty;
        attributeSourceText = string.Empty;
        attributeSpan = null;

        switch (inline) {
            case LinkInline link when link.LabelInlines != null:
                return TryConsumeLinkLabelTrailingAttributes(
                    link,
                    out replacement,
                    out attributes,
                    out attributeSourceText,
                    out attributeSpan);

            case IInlineContainerMarkdownInline container when container.NestedInlines != null:
                return TryConsumeNestedInlineSequenceTrailingAttributes(
                    container.NestedInlines,
                    out attributes,
                    out attributeSourceText,
                    out attributeSpan);

            case ImageInline image:
                return TryConsumeImageAltTrailingAttributes(
                    image,
                    out replacement,
                    out attributes,
                    out attributeSourceText,
                    out attributeSpan);

            case ImageLinkInline imageLink:
                return TryConsumeImageLinkAltTrailingAttributes(
                    imageLink,
                    out replacement,
                    out attributes,
                    out attributeSourceText,
                    out attributeSpan);
        }

        return false;
    }

    private static bool TryConsumeLinkLabelTrailingAttributes(
        LinkInline link,
        out IMarkdownInline? replacement,
        out MarkdownAttributeSet attributes,
        out string attributeSourceText,
        out MarkdownSourceSpan? attributeSpan) {
        replacement = null;
        attributes = MarkdownAttributeSet.Empty;
        attributeSourceText = string.Empty;
        attributeSpan = null;

        if (link == null ||
            link.LabelInlines == null ||
            !TryConsumeNestedInlineSequenceTrailingAttributes(
                link.LabelInlines,
                out attributes,
                out attributeSourceText,
                out attributeSpan)) {
            return false;
        }

        var rewritten = new LinkInline(link.LabelInlines, link.Url, link.Title, link.LinkTarget, link.LinkRel);
        CopyLinkInlineSourceMetadata(link, rewritten);
        replacement = rewritten;
        return true;
    }

    private static bool TryConsumeNestedInlineSequenceTrailingAttributes(
        InlineSequence sequence,
        out MarkdownAttributeSet attributes,
        out string attributeSourceText,
        out MarkdownSourceSpan? attributeSpan) {
        attributes = MarkdownAttributeSet.Empty;
        attributeSourceText = string.Empty;
        attributeSpan = null;

        if (sequence == null || sequence.Nodes.Count == 0) {
            return false;
        }

        var lastIndex = sequence.Nodes.Count - 1;
        if (sequence.Nodes[lastIndex] is IInlineContainerMarkdownInline container &&
            container.NestedInlines != null) {
            if (TryConsumeContainerGenericAttributes(
                container,
                out attributes,
                out attributeSourceText,
                out attributeSpan)) {
                return true;
            }

            if (TryConsumeNestedInlineSequenceTrailingAttributes(
                    container.NestedInlines,
                    out attributes,
                    out attributeSourceText,
                    out attributeSpan)) {
                return true;
            }
        }

        if (sequence.Nodes[lastIndex] is not TextRun textRun) {
            return false;
        }

        var allowEmptyText = lastIndex > 0;
        if (!TryConsumeTrailingAttributesFromText(
                textRun.Text,
                MarkdownInlineSourceSpans.Get(textRun),
                out var textWithoutAttributeBlock,
                out attributes,
                out attributeSourceText,
                out attributeSpan,
                out var remainingTextSpan,
                allowEmptyText)) {
            return false;
        }

        var replacement = new List<IMarkdownInline>(sequence.Nodes.Count);
        for (var i = 0; i < lastIndex; i++) {
            replacement.Add(sequence.Nodes[i]);
        }

        if (!string.IsNullOrEmpty(textWithoutAttributeBlock)) {
            var remainingText = new TextRun(textWithoutAttributeBlock);
            MarkdownInlineSourceSpans.Set(remainingText, remainingTextSpan);
            replacement.Add(remainingText);
        }

        sequence.ReplaceItems(replacement);
        return true;
    }

    private static bool TryConsumeContainerGenericAttributes(
        IInlineContainerMarkdownInline container,
        out MarkdownAttributeSet attributes,
        out string attributeSourceText,
        out MarkdownSourceSpan? attributeSpan) {
        attributes = MarkdownAttributeSet.Empty;
        attributeSourceText = string.Empty;
        attributeSpan = null;

        if (container is not MarkdownObject markdownObject || markdownObject.Attributes.IsEmpty) {
            return false;
        }

        attributes = markdownObject.Attributes;
        attributeSourceText = MarkdownGenericAttributeSourceSpans.GetSourceText(markdownObject) ?? string.Empty;
        attributeSpan = MarkdownGenericAttributeSourceSpans.GetSourceSpan(markdownObject);
        markdownObject.SetAttributes(MarkdownAttributeSet.Empty);
        return !attributes.IsEmpty;
    }

    private static bool TryConsumeImageAltTrailingAttributes(
        ImageInline image,
        out IMarkdownInline? replacement,
        out MarkdownAttributeSet attributes,
        out string attributeSourceText,
        out MarkdownSourceSpan? attributeSpan) {
        replacement = null;
        attributes = MarkdownAttributeSet.Empty;
        attributeSourceText = string.Empty;
        attributeSpan = null;

        if (image == null ||
            !TryConsumeTrailingAttributesFromText(
                image.Alt,
                MarkdownInlineMetadataSourceSpans.GetImageAltSpan(image),
                out var altWithoutAttributeBlock,
                out attributes,
                out attributeSourceText,
                out attributeSpan,
                out var remainingAltSpan)) {
            return false;
        }

        var plainAlt = TryStripPlainAltTrailingAttributes(image.PlainAlt, out var strippedPlainAlt)
            ? strippedPlainAlt
            : altWithoutAttributeBlock;
        var rewritten = new ImageInline(altWithoutAttributeBlock, image.Src, image.Title, plainAlt);
        CopyImageInlineSourceMetadata(image, rewritten, remainingAltSpan);
        replacement = rewritten;
        return true;
    }

    private static bool TryConsumeImageLinkAltTrailingAttributes(
        ImageLinkInline image,
        out IMarkdownInline? replacement,
        out MarkdownAttributeSet attributes,
        out string attributeSourceText,
        out MarkdownSourceSpan? attributeSpan) {
        replacement = null;
        attributes = MarkdownAttributeSet.Empty;
        attributeSourceText = string.Empty;
        attributeSpan = null;

        if (image == null ||
            !TryConsumeTrailingAttributesFromText(
                image.Alt,
                MarkdownInlineMetadataSourceSpans.GetImageAltSpan(image),
                out var altWithoutAttributeBlock,
                out attributes,
                out attributeSourceText,
                out attributeSpan,
                out var remainingAltSpan)) {
            return false;
        }

        var plainAlt = TryStripPlainAltTrailingAttributes(image.PlainAlt, out var strippedPlainAlt)
            ? strippedPlainAlt
            : altWithoutAttributeBlock;
        var rewritten = new ImageLinkInline(altWithoutAttributeBlock, image.ImageUrl, image.LinkUrl, image.Title, image.LinkTitle, plainAlt);
        CopyImageLinkInlineSourceMetadata(image, rewritten, remainingAltSpan);
        replacement = rewritten;
        return true;
    }

    private static bool TryConsumeTrailingAttributesFromText(
        string value,
        MarkdownSourceSpan? valueSourceSpan,
        out string textWithoutAttributeBlock,
        out MarkdownAttributeSet attributes,
        out string attributeSourceText,
        out MarkdownSourceSpan? attributeSpan,
        out MarkdownSourceSpan? remainingTextSpan,
        bool allowEmptyText = false) {
        attributeSourceText = string.Empty;
        attributeSpan = null;
        remainingTextSpan = null;

        if (!MarkdownGenericAttributeParser.TryConsumeTrailingAttributeBlock(
            value,
            out textWithoutAttributeBlock,
            out attributes,
            out var attributeStart,
            out var attributeEnd,
            requireLeadingWhitespace: false)) {
            return false;
        }

        if (string.IsNullOrEmpty(textWithoutAttributeBlock) && !allowEmptyText) {
            attributes = MarkdownAttributeSet.Empty;
            return false;
        }

        attributeSourceText = value.Substring(attributeStart, attributeEnd - attributeStart + 1);
        attributeSpan = SliceSourceSpan(valueSourceSpan, value, attributeStart, attributeSourceText.Length);
        remainingTextSpan = TrimSourceSpanToConsumedPrefix(valueSourceSpan, value, textWithoutAttributeBlock.Length);
        return !attributes.IsEmpty;
    }

    private static bool TryStripPlainAltTrailingAttributes(string value, out string stripped) {
        if (MarkdownGenericAttributeParser.TryConsumeTrailingAttributeBlock(
            value,
            out stripped,
            out var attributes,
            out _,
            out _,
            requireLeadingWhitespace: false)) {
            return !attributes.IsEmpty;
        }

        stripped = value ?? string.Empty;
        return false;
    }

    private static MarkdownSourceSpan? SliceSourceSpan(MarkdownSourceSpan? span, string? text, int startIndex, int length) {
        if (!span.HasValue || length <= 0) {
            return null;
        }

        var value = span.Value;
        if (value.StartLine != value.EndLine || !value.StartColumn.HasValue) {
            return null;
        }

        var startColumn = AdvanceSourceColumn(value.StartColumn.Value, text, startIndex);
        var endColumn = AdvanceSourceColumn(value.StartColumn.Value, text, startIndex + length) - 1;
        int? startOffset = value.StartOffset.HasValue ? value.StartOffset.Value + startIndex : null;
        int? endOffset = startOffset.HasValue ? startOffset.Value + length - 1 : value.EndOffset;
        return new MarkdownSourceSpan(
            value.StartLine,
            startColumn,
            value.EndLine,
            endColumn,
            startOffset,
            endOffset);
    }

    private static MarkdownSourceSpan? TrimSourceSpanToConsumedPrefix(MarkdownSourceSpan? span, string? text, int consumedLength) =>
        span.HasValue ? TrimSourceSpanToConsumedPrefix(span.Value, text, consumedLength) : null;

    private static void CopyLinkInlineSourceMetadata(LinkInline source, LinkInline target) {
        MarkdownInlineSourceSpans.Set(target, MarkdownInlineSourceSpans.Get(source));
        MarkdownInlineMetadataSourceSpans.SetLinkParts(
            target,
            MarkdownInlineMetadataSourceSpans.GetLinkTargetSpan(source),
            MarkdownInlineMetadataSourceSpans.GetLinkTitleSpan(source),
            MarkdownInlineMetadataSourceSpans.GetLinkHtmlTargetSpan(source),
            MarkdownInlineMetadataSourceSpans.GetLinkHtmlRelSpan(source),
            MarkdownInlineMetadataSourceSpans.GetAutolinkLiteral(source));
        CopyFormattingMarkerMetadata(source, target);
    }

    private static void CopyImageInlineSourceMetadata(ImageInline source, ImageInline target, MarkdownSourceSpan? altSpan) {
        MarkdownInlineSourceSpans.Set(target, MarkdownInlineSourceSpans.Get(source));
        MarkdownInlineMetadataSourceSpans.SetImageParts(
            target,
            altSpan,
            MarkdownInlineMetadataSourceSpans.GetImageSourceSpan(source),
            MarkdownInlineMetadataSourceSpans.GetImageTitleSpan(source));
        CopyFormattingMarkerMetadata(source, target);
    }

    private static void CopyImageLinkInlineSourceMetadata(ImageLinkInline source, ImageLinkInline target, MarkdownSourceSpan? altSpan) {
        MarkdownInlineSourceSpans.Set(target, MarkdownInlineSourceSpans.Get(source));
        MarkdownInlineMetadataSourceSpans.SetImageLinkParts(
            target,
            altSpan,
            MarkdownInlineMetadataSourceSpans.GetImageSourceSpan(source),
            MarkdownInlineMetadataSourceSpans.GetImageTitleSpan(source),
            MarkdownInlineMetadataSourceSpans.GetImageLinkTargetSpan(source),
            MarkdownInlineMetadataSourceSpans.GetImageLinkTitleSpan(source));
        CopyFormattingMarkerMetadata(source, target);
    }

    private static void CopyFormattingMarkerMetadata(MarkdownInline source, MarkdownInline target) {
        MarkdownInlineMetadataSourceSpans.SetFormattingMarkers(
            target,
            MarkdownInlineMetadataSourceSpans.GetOpeningMarker(source) ?? string.Empty,
            MarkdownInlineMetadataSourceSpans.GetOpeningMarkerSpan(source),
            MarkdownInlineMetadataSourceSpans.GetClosingMarker(source) ?? string.Empty,
            MarkdownInlineMetadataSourceSpans.GetClosingMarkerSpan(source),
            MarkdownInlineMetadataSourceSpans.GetSeparatorMarker(source),
            MarkdownInlineMetadataSourceSpans.GetSeparatorMarkerSpan(source));
    }
}

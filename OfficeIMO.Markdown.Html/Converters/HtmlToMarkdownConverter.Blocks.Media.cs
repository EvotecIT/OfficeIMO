using AngleSharp.Dom;
using OfficeIMO.Html;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

internal sealed partial class HtmlToMarkdownConverter {
    private static IEnumerable<IMarkdownBlock> ConvertMediaElement(IElement element, ConversionContext context) {
        if (!context.Options.PreserveUnsupportedBlocks) {
            return ConvertUnknownElementChildrenToBlocks(element, context);
        }

        IElement clone = (IElement)element.Clone(deep: true);
        ResolveMediaUrlAttribute(clone, "src", context);
        ResolveMediaUrlAttribute(clone, "poster", context);
        foreach (IElement child in clone.QuerySelectorAll("source[src], track[src]")) {
            ResolveMediaUrlAttribute(child, "src", context);
        }
        return new IMarkdownBlock[] { new HtmlRawBlock(NormalizeRawElement(clone, context)) };
    }

    private static void ResolveMediaUrlAttribute(IElement element, string attributeName, ConversionContext context) {
        string? value = element.GetAttribute(attributeName);
        if (string.IsNullOrWhiteSpace(value)) return;
        string resolved = ResolveResourceUrl(value, context);
        if (resolved.Length == 0) {
            element.RemoveAttribute(attributeName);
        } else {
            element.SetAttribute(attributeName, resolved);
        }
    }

    private static IEnumerable<IMarkdownBlock> ConvertImageElement(IElement element, ConversionContext context) {
        if (!TryCreateImageBlock(element, context, out var image)) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { image };
    }

    private static IEnumerable<IMarkdownBlock> ConvertPictureElement(IElement element, ConversionContext context) {
        if (element == null) {
            return Array.Empty<IMarkdownBlock>();
        }

        string preferredSrc = ResolvePictureSource(element, context);
        bool hasUnsafePictureCandidate = HasRejectedPictureSourceCandidate(element, context)
                                         || (context.Options.Base64Images != HtmlBase64ImageHandling.Include
                                             && HasBase64PictureCandidate(element, context));
        var imageElement = FindFirstDescendantByEffectiveTagName(element, context, "IMG");
        if (imageElement != null && TryCreateImageBlock(imageElement, context, out var imageBlock)) {
            if (!string.IsNullOrWhiteSpace(preferredSrc) && !TryApplyBase64ImageHandling(ref preferredSrc, context)) {
                preferredSrc = string.Empty;
            }

            imageBlock = !string.IsNullOrWhiteSpace(preferredSrc)
                ? CreateImageBlock(preferredSrc, imageElement, element, context)
                : CreateImageBlock(imageBlock.Path, imageElement, element, context);

            return new IMarkdownBlock[] { imageBlock };
        }

        if (!string.IsNullOrWhiteSpace(preferredSrc) && !TryApplyBase64ImageHandling(ref preferredSrc, context)) {
            preferredSrc = string.Empty;
        }

        if (string.IsNullOrWhiteSpace(preferredSrc)) {
            if (hasUnsafePictureCandidate) {
                return Array.Empty<IMarkdownBlock>();
            }

            return context.Options.PreserveUnsupportedBlocks
                ? new IMarkdownBlock[] { new HtmlRawBlock(NormalizeRawElement(element, context)) }
                : Array.Empty<IMarkdownBlock>();
        }

        var pictureImage = new ImageBlock(preferredSrc, alt: null, title: null);
        ApplyPictureMetadata(element, pictureImage, null, context);
        return new IMarkdownBlock[] { pictureImage };
    }

    private static IEnumerable<IMarkdownBlock> ConvertFigureElement(IElement element, ConversionContext context) {
        var directCaption = FindDirectChildByEffectiveTagName(element, context, "FIGCAPTION");
        var directMediaContainer = element.Children.FirstOrDefault(child => TryResolveFigureMediaElement(child, context, out _));

        if (directMediaContainer != null && TryResolveFigureMediaElement(directMediaContainer, context, out var directMedia)) {
            var figureBlocks = new List<IMarkdownBlock>();
            foreach (var child in element.ChildNodes) {
                if (ReferenceEquals(child, directCaption)) {
                    continue;
                }

                if (ReferenceEquals(child, directMediaContainer)) {
                    var mediaBlocks = ConvertFigureMediaElement(directMedia, context);
                    ApplyFigureCaptionToMedia(mediaBlocks, directCaption);
                    figureBlocks.AddRange(mediaBlocks);
                    continue;
                }

                figureBlocks.AddRange(ConvertNodesToBlocks(new[] { child }, context));
            }

            if (figureBlocks.Count > 0) {
                return figureBlocks;
            }
        }

        var imageElement = FindFirstDescendantByEffectiveTagName(element, context, "IMG");
        if (imageElement == null) {
            var pictureElement = FindFirstDescendantByEffectiveTagName(element, context, "PICTURE");
            if (pictureElement != null) {
                var pictureBlocks = ConvertPictureElement(pictureElement, context).ToList();
                if (pictureBlocks.Count > 0) {
                    ApplyFigureCaptionToMedia(pictureBlocks, directCaption ?? FindFirstDescendantByEffectiveTagName(element, context, "FIGCAPTION"));

                    return pictureBlocks;
                }
            }

            if (HasDirectBlockChildren(element, context)) {
                return ConvertNodesToBlocks(element.ChildNodes, context);
            }

            var inlineSequence = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(element.ChildNodes, context));
            if (!HasVisibleInlineContent(inlineSequence)) {
                return Array.Empty<IMarkdownBlock>();
            }

            return new IMarkdownBlock[] { new ParagraphBlock(inlineSequence) };
        }

        var blocks = ConvertImageElement(imageElement, context).ToList();
        ApplyFigureCaptionToMedia(blocks, directCaption ?? FindFirstDescendantByEffectiveTagName(element, context, "FIGCAPTION"));

        return blocks;
    }

    private static void ApplyFigureCaptionToMedia(IReadOnlyList<IMarkdownBlock> blocks, IElement? captionElement) {
        if (captionElement == null || blocks == null || blocks.Count != 1 || blocks[0] is not ImageBlock imageBlock) {
            return;
        }

        imageBlock.Caption = NormalizeBlockText(captionElement.TextContent);
    }

    private static bool IsLinkedFigureMediaAnchor(IElement element, ConversionContext context) {
        return TryResolveAnchorMediaElement(element, context, out _);
    }

    private static bool TryResolveFigureMediaElement(IElement element, ConversionContext context, out IElement mediaElement) {
        return TryResolvePureWrapperElement(
            element,
            candidate =>
                HasEffectiveTagName(candidate, context, "IMG")
                || HasEffectiveTagName(candidate, context, "PICTURE")
                || IsLinkedFigureMediaAnchor(candidate, context),
            candidate => !HasEffectiveTagName(candidate, context, "FIGCAPTION"),
            out mediaElement);
    }

    private static bool TryResolveAnchorMediaElement(IElement element, ConversionContext context, out IElement mediaElement) {
        mediaElement = null!;
        if (element == null || !HasEffectiveTagName(element, context, "A") || HasVisibleOwnTextNodes(element)) {
            return false;
        }

        foreach (var childNode in element.ChildNodes) {
            switch (childNode) {
                case IComment:
                    continue;
                case IText textNode when string.IsNullOrWhiteSpace(textNode.Data):
                    continue;
                case IElement childElement when IsIgnorableMediaWrapperChild(childElement):
                    continue;
                case IElement childElement when HasEffectiveTagName(childElement, context, "IMG")
                    || HasEffectiveTagName(childElement, context, "PICTURE"):
                    mediaElement = childElement;
                    return true;
                case IElement childElement when !HasEffectiveTagName(childElement, context, "A")
                    && TryResolvePureWrapperElement(
                        childElement,
                        candidate =>
                            HasEffectiveTagName(candidate, context, "IMG")
                            || HasEffectiveTagName(candidate, context, "PICTURE"),
                        candidate => !HasEffectiveTagName(candidate, context, "A"),
                        out mediaElement):
                    return true;
                default:
                    return false;
            }
        }

        return false;
    }

    private static bool IsIgnorableMediaWrapperChild(IElement element) {
        if (element == null) {
            return false;
        }

        return element.TagName.Equals("NOSCRIPT", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("SCRIPT", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("STYLE", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("TEMPLATE", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryResolvePureWrapperElement(
        IElement element,
        Func<IElement, bool> terminalPredicate,
        Func<IElement, bool> canRecursePredicate,
        out IElement resolvedElement) {
        resolvedElement = null!;
        if (element == null) {
            return false;
        }

        if (terminalPredicate(element)) {
            resolvedElement = element;
            return true;
        }

        if (!canRecursePredicate(element) || HasVisibleOwnTextNodes(element)) {
            return false;
        }

        IElement? onlyChildElement = null;
        foreach (var childNode in element.ChildNodes) {
            switch (childNode) {
                case IComment:
                    continue;
                case IText textNode when string.IsNullOrWhiteSpace(textNode.Data):
                    continue;
                case IElement childElement when IsIgnorableMediaWrapperChild(childElement):
                    continue;
                case IElement childElement:
                    if (onlyChildElement != null) {
                        return false;
                    }

                    onlyChildElement = childElement;
                    break;
                default:
                    return false;
            }
        }

        return onlyChildElement != null
            && TryResolvePureWrapperElement(onlyChildElement, terminalPredicate, canRecursePredicate, out resolvedElement);
    }

    private static bool HasVisibleOwnTextNodes(IElement element) {
        foreach (var childNode in element.ChildNodes) {
            if (childNode is IText textNode && !string.IsNullOrWhiteSpace(textNode.Data)) {
                return true;
            }
        }

        return false;
    }

    private static List<IMarkdownBlock> ConvertFigureMediaElement(IElement element, ConversionContext context) {
        if (HasEffectiveTagName(element, context, "PICTURE")) {
            return ConvertPictureElement(element, context).ToList();
        }

        if (HasEffectiveTagName(element, context, "IMG")) {
            return ConvertImageElement(element, context).ToList();
        }

        if (TryCreateLinkedImageBlockFromAnchor(element, context, out var linkedImage)) {
            return new List<IMarkdownBlock> { linkedImage };
        }

        return ConvertNodesToBlocks(new[] { element }, context);
    }

    private static bool HasBase64PictureCandidate(IElement element, ConversionContext context) {
        if (element == null || context == null) {
            return false;
        }

        foreach (var child in element.Children) {
            if (!HasEffectiveTagName(child, context, "SOURCE")) {
                continue;
            }

            if (HasBase64SrcSetAttribute(child, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset")
                || HasBase64UrlAttribute(child, "src", "data-src", "data-original-src", "data-lazy-src")) {
                return true;
            }
        }

        var imageElement = FindFirstDescendantByEffectiveTagName(element, context, "IMG");
        if (imageElement == null) {
            return false;
        }

        return HasBase64SrcSetAttribute(imageElement, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset")
               || HasBase64UrlAttribute(imageElement, "src", "data-src", "data-original", "data-original-src", "data-lazy-src");
    }

    private static bool HasRejectedPictureSourceCandidate(IElement element, ConversionContext context) {
        if (element == null || context == null) {
            return false;
        }

        foreach (var child in element.Children) {
            if (!HasEffectiveTagName(child, context, "SOURCE")) {
                continue;
            }

            if (HasRejectedSrcSetAttribute(child, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset")
                || HasRejectedUrlAttribute(child, context, "src", "data-src", "data-original-src", "data-lazy-src")) {
                return true;
            }
        }

        var imageElement = FindFirstDescendantByEffectiveTagName(element, context, "IMG");
        if (imageElement == null) {
            return false;
        }

        return HasRejectedSrcSetAttribute(imageElement, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset")
               || HasRejectedUrlAttribute(imageElement, context, "src", "data-src", "data-original", "data-original-src", "data-lazy-src");
    }

    private static bool HasRejectedMediaSourceCandidate(IElement element, ConversionContext context) {
        if (element == null || context == null) {
            return false;
        }

        if (HasEffectiveTagName(element, context, "PICTURE")) {
            return HasRejectedPictureSourceCandidate(element, context);
        }

        return HasEffectiveTagName(element, context, "IMG")
               && (HasRejectedSrcSetAttribute(element, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset")
                   || HasRejectedUrlAttribute(element, context, "src", "data-src", "data-original", "data-original-src", "data-lazy-src"));
    }

    private static bool HasBlockedBase64MediaSourceCandidate(IElement element, ConversionContext context) {
        if (element == null || context == null || context.Options.Base64Images == HtmlBase64ImageHandling.Include) {
            return false;
        }

        if (HasEffectiveTagName(element, context, "PICTURE")) {
            return HasBlockedBase64PictureCandidate(element, context);
        }

        return HasEffectiveTagName(element, context, "IMG")
               && (HasBlockedBase64SrcSetAttribute(element, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset")
                   || HasBlockedBase64UrlAttribute(element, context, "src", "data-src", "data-original", "data-original-src", "data-lazy-src"));
    }

    private static bool HasBlockedBase64PictureCandidate(IElement element, ConversionContext context) {
        foreach (var child in element.Children) {
            if (!HasEffectiveTagName(child, context, "SOURCE")) {
                continue;
            }

            if (HasBlockedBase64SrcSetAttribute(child, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset")
                || HasBlockedBase64UrlAttribute(child, context, "src", "data-src", "data-original-src", "data-lazy-src")) {
                return true;
            }
        }

        var imageElement = FindFirstDescendantByEffectiveTagName(element, context, "IMG");
        if (imageElement == null) {
            return false;
        }

        return HasBlockedBase64SrcSetAttribute(imageElement, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset")
               || HasBlockedBase64UrlAttribute(imageElement, context, "src", "data-src", "data-original", "data-original-src", "data-lazy-src");
    }

    private static bool HasBase64UrlAttribute(IElement element, params string[] attributeNames) {
        if (element == null || attributeNames == null) {
            return false;
        }

        foreach (string attributeName in attributeNames) {
            if (IsBase64ImageDataUri(element.GetAttribute(attributeName))) {
                return true;
            }
        }

        return false;
    }

    private static bool HasBlockedBase64UrlAttribute(IElement element, ConversionContext context, params string[] attributeNames) {
        if (element == null || context == null || attributeNames == null) {
            return false;
        }

        foreach (string attributeName in attributeNames) {
            if (IsBlockedBase64ImageSource(element.GetAttribute(attributeName), context)) {
                return true;
            }
        }

        return false;
    }

    private static bool HasRejectedUrlAttribute(IElement element, ConversionContext context, params string[] attributeNames) {
        if (element == null || context == null || attributeNames == null) {
            return false;
        }

        foreach (string attributeName in attributeNames) {
            if (IsRejectedImageSource(element.GetAttribute(attributeName), context)) {
                return true;
            }
        }

        return false;
    }

    private static bool HasBase64SrcSetAttribute(IElement element, params string[] attributeNames) {
        if (element == null || attributeNames == null) {
            return false;
        }

        foreach (string attributeName in attributeNames) {
            foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(element.GetAttribute(attributeName))) {
                if (IsBase64ImageDataUri(candidate.Url)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool HasRejectedSrcSetAttribute(IElement element, ConversionContext context, params string[] attributeNames) {
        if (element == null || context == null || attributeNames == null) {
            return false;
        }

        foreach (string attributeName in attributeNames) {
            foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(element.GetAttribute(attributeName))) {
                if (IsRejectedImageSource(candidate.Url, context)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool HasBlockedBase64SrcSetAttribute(IElement element, ConversionContext context, params string[] attributeNames) {
        if (element == null || context == null || attributeNames == null) {
            return false;
        }

        foreach (string attributeName in attributeNames) {
            foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(element.GetAttribute(attributeName))) {
                if (IsBlockedBase64ImageSource(candidate.Url, context)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool IsRejectedImageSource(string? source, ConversionContext context) {
        if (string.IsNullOrWhiteSpace(source)) {
            return false;
        }

        return string.IsNullOrWhiteSpace(HtmlUrlPolicyEvaluator.ResolveUrl(
            source,
            context.Options.BaseUri,
            context.Options.ResourceUrlPolicy ?? context.Options.UrlPolicy));
    }

    private static bool IsBlockedBase64ImageSource(string? source, ConversionContext context) {
        if (string.IsNullOrWhiteSpace(source) || !IsBase64ImageDataUri(source)) {
            return false;
        }

        return !CanUseImageCandidateWithoutSideEffects(source, context);
    }

}

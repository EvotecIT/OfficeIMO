using AngleSharp.Dom;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

public sealed partial class HtmlToMarkdownConverter {
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
        var imageElement = element.QuerySelector("img");
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
            return context.Options.PreserveUnsupportedBlocks
                ? new IMarkdownBlock[] { new HtmlRawBlock(element.OuterHtml) }
                : Array.Empty<IMarkdownBlock>();
        }

        var pictureImage = new ImageBlock(preferredSrc, alt: null, title: null);
        ApplyPictureMetadata(element, pictureImage, null, context);
        return new IMarkdownBlock[] { pictureImage };
    }

    private static IEnumerable<IMarkdownBlock> ConvertFigureElement(IElement element, ConversionContext context) {
        var directCaption = element.Children.FirstOrDefault(child => child.TagName.Equals("FIGCAPTION", StringComparison.OrdinalIgnoreCase));
        var directMediaContainer = element.Children.FirstOrDefault(child => TryResolveFigureMediaElement(child, out _));

        if (directMediaContainer != null && TryResolveFigureMediaElement(directMediaContainer, out var directMedia)) {
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

        var imageElement = element.QuerySelector("img");
        if (imageElement == null) {
            var pictureElement = element.QuerySelector("picture");
            if (pictureElement != null) {
                var pictureBlocks = ConvertPictureElement(pictureElement, context).ToList();
                if (pictureBlocks.Count > 0) {
                    ApplyFigureCaptionToMedia(pictureBlocks, directCaption ?? element.QuerySelector("figcaption"));

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
        ApplyFigureCaptionToMedia(blocks, directCaption ?? element.QuerySelector("figcaption"));

        return blocks;
    }

    private static void ApplyFigureCaptionToMedia(IReadOnlyList<IMarkdownBlock> blocks, IElement? captionElement) {
        if (captionElement == null || blocks == null || blocks.Count != 1 || blocks[0] is not ImageBlock imageBlock) {
            return;
        }

        imageBlock.Caption = NormalizeBlockText(captionElement.TextContent);
    }

    private static bool IsLinkedFigureMediaAnchor(IElement element) {
        return TryResolveAnchorMediaElement(element, out _);
    }

    private static bool TryResolveFigureMediaElement(IElement element, out IElement mediaElement) {
        return TryResolvePureWrapperElement(
            element,
            candidate =>
                candidate.TagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)
                || candidate.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)
                || IsLinkedFigureMediaAnchor(candidate),
            candidate => !candidate.TagName.Equals("FIGCAPTION", StringComparison.OrdinalIgnoreCase),
            out mediaElement);
    }

    private static bool TryResolveAnchorMediaElement(IElement element, out IElement mediaElement) {
        mediaElement = null!;
        if (element == null || !element.TagName.Equals("A", StringComparison.OrdinalIgnoreCase) || HasVisibleOwnTextNodes(element)) {
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
                case IElement childElement when childElement.TagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)
                    || childElement.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase):
                    mediaElement = childElement;
                    return true;
                case IElement childElement when !childElement.TagName.Equals("A", StringComparison.OrdinalIgnoreCase)
                    && TryResolvePureWrapperElement(
                        childElement,
                        candidate =>
                            candidate.TagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)
                            || candidate.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase),
                        candidate => !candidate.TagName.Equals("A", StringComparison.OrdinalIgnoreCase),
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
        if (element.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)) {
            return ConvertPictureElement(element, context).ToList();
        }

        if (element.TagName.Equals("IMG", StringComparison.OrdinalIgnoreCase)) {
            return ConvertImageElement(element, context).ToList();
        }

        if (TryCreateLinkedImageBlockFromAnchor(element, context, out var linkedImage)) {
            return new List<IMarkdownBlock> { linkedImage };
        }

        return ConvertNodesToBlocks(new[] { element }, context);
    }

}

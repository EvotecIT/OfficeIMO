using AngleSharp.Dom;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using System.Text;

namespace OfficeIMO.Markdown.Html;

internal sealed partial class HtmlToMarkdownConverter {
    private static string ResolveImageSource(IElement element, ConversionContext context) {
        return HtmlImageSourceResolver.ResolveImageSource(
            element,
            context.Options.BaseUri,
            context.Options.ResourceUrlPolicy ?? context.Options.UrlPolicy,
            allowParentPictureFallback: true);
    }

    private static IReadOnlyList<string> ResolveImageSourceCandidates(IElement element, ConversionContext context) {
        return HtmlImageSourceResolver.ResolveImageSourceCandidates(
            element,
            context.Options.BaseUri,
            context.Options.ResourceUrlPolicy ?? context.Options.UrlPolicy,
            allowParentPictureFallback: true);
    }

    private static string ResolveDirectImageSource(IElement element, ConversionContext context) {
        return HtmlImageSourceResolver.ResolveImageSource(
            element,
            context.Options.BaseUri,
            context.Options.ResourceUrlPolicy ?? context.Options.UrlPolicy,
            allowParentPictureFallback: false);
    }

    private static string ResolvePictureSource(IElement pictureElement, ConversionContext context) {
        if (TryResolveUsableImageCandidate(ResolvePictureSourceCandidates(pictureElement, context), context, out string source)) {
            return source;
        }

        return string.Empty;
    }

    private static IReadOnlyList<string> ResolvePictureSourceCandidates(IElement pictureElement, ConversionContext context) {
        var candidates = new List<string>();
        if (pictureElement == null || context == null) {
            return candidates;
        }

        foreach (var child in pictureElement.Children) {
            if (!HasEffectiveTagName(child, context, "SOURCE")) {
                continue;
            }

            AddResolvedSrcSetCandidateAttributes(candidates, child, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset");
            AddResolvedCandidate(candidates, ResolveUrlAttributes(child, context, "src", "data-src", "data-original-src", "data-lazy-src"));
        }

        return candidates;
    }

    private static void AddResolvedSrcSetCandidateAttributes(IList<string> candidates, IElement element, ConversionContext context, params string[] attributeNames) {
        if (candidates == null || element == null || context == null || attributeNames == null) {
            return;
        }

        for (int i = 0; i < attributeNames.Length; i++) {
            IReadOnlyList<HtmlSrcSetCandidate> srcSetCandidates = HtmlImageSourceResolver.ResolveSrcSetCandidates(
                element.GetAttribute(attributeNames[i]),
                context.Options.BaseUri,
                context.Options.ResourceUrlPolicy ?? context.Options.UrlPolicy);

            for (int candidateIndex = 0; candidateIndex < srcSetCandidates.Count; candidateIndex++) {
                AddResolvedCandidate(candidates, srcSetCandidates[candidateIndex].Url);
            }
        }
    }

    private static void AddResolvedCandidate(IList<string> candidates, string? value) {
        if (candidates == null || string.IsNullOrWhiteSpace(value)) {
            return;
        }

        candidates.Add(value!);
    }

    private static string ResolveUrlFromSrcSet(string? rawSrcSet, ConversionContext context) {
        return HtmlImageSourceResolver.ResolveUrlFromSrcSet(rawSrcSet, context.Options.BaseUri, context.Options.ResourceUrlPolicy ?? context.Options.UrlPolicy);
    }

    private static string ResolveNormalizedSrcSet(string? rawSrcSet, ConversionContext context) {
        return HtmlImageSourceResolver.ResolveNormalizedSrcSet(rawSrcSet, context.Options.BaseUri, context.Options.ResourceUrlPolicy ?? context.Options.UrlPolicy);
    }

    private static string ResolveUrlFromSrcSetAttributes(IElement element, ConversionContext context, params string[] attributeNames) {
        return HtmlImageSourceResolver.ResolveUrlFromSrcSetAttributes(element, context.Options.BaseUri, context.Options.ResourceUrlPolicy ?? context.Options.UrlPolicy, attributeNames);
    }

    private static string ResolveNormalizedSrcSetAttributes(IElement element, ConversionContext context, params string[] attributeNames) {
        return HtmlImageSourceResolver.ResolveNormalizedSrcSetAttributes(element, context.Options.BaseUri, context.Options.ResourceUrlPolicy ?? context.Options.UrlPolicy, attributeNames);
    }

    private static string ResolveUrlAttributes(IElement element, ConversionContext context, params string[] attributeNames) {
        return HtmlImageSourceResolver.ResolveUrlAttributes(element, context.Options.BaseUri, context.Options.ResourceUrlPolicy ?? context.Options.UrlPolicy, attributeNames);
    }

    private static void ApplyImageDimensions(IElement element, ImageBlock image) {
        if (TryParseImageDimension(element.GetAttribute("width"), out double width)
            || TryParseStyleDimension(element.GetAttribute("style"), "width", out width)) {
            image.Width = width;
        }

        if (TryParseImageDimension(element.GetAttribute("height"), out double height)
            || TryParseStyleDimension(element.GetAttribute("style"), "height", out height)) {
            image.Height = height;
        }
    }

    private static bool TryCreateImageBlock(IElement element, ConversionContext context, out ImageBlock image) {
        image = null!;
        IReadOnlyList<string> candidates = ResolveImageSourceCandidates(element, context);
        string firstCandidate = candidates.Count == 0 ? string.Empty : candidates[0];
        if ((candidates.Count == 0 || IsLikelyPlaceholderImageSource(firstCandidate))
            && TryCreateImageBlockFromNoscriptFallback(element, context, out image)) {
            return true;
        }

        if (TryResolveUsableImageCandidate(candidates, context, out string src)) {
            image = CreateImageBlock(src, element);
            return true;
        }

        return false;
    }

    private static bool CanCreateImageBlockWithoutSideEffects(IElement element, ConversionContext context) {
        IReadOnlyList<string> candidates = ResolveImageSourceCandidates(element, context);
        string firstCandidate = candidates.Count == 0 ? string.Empty : candidates[0];
        if ((candidates.Count == 0 || IsLikelyPlaceholderImageSource(firstCandidate))
            && CanCreateImageBlockFromNoscriptFallbackWithoutSideEffects(element, context)) {
            return true;
        }

        return HasUsableImageCandidateWithoutSideEffects(candidates, context);
    }

    private static bool CanCreateImageBlockFromNoscriptFallbackWithoutSideEffects(IElement element, ConversionContext context) {
        if (!TryResolveAssociatedNoscriptMediaElement(element, out var fallbackMediaElement)) {
            return false;
        }

        if (HasEffectiveTagName(fallbackMediaElement, context, "PICTURE")) {
            return CanCreatePictureImageBlockWithoutSideEffects(fallbackMediaElement, context);
        }

        return HasUsableImageCandidateWithoutSideEffects(ResolveImageSourceCandidates(fallbackMediaElement, context), context);
    }

    private static bool CanCreatePictureImageBlockWithoutSideEffects(IElement element, ConversionContext context) {
        if (element == null || context == null) {
            return false;
        }

        if (HasUsableImageCandidateWithoutSideEffects(
            ResolvePictureSourceCandidates(element, context),
            context)) {
            return true;
        }

        var imageElement = FindFirstDescendantByEffectiveTagName(element, context, "IMG");
        return imageElement != null && CanCreateImageBlockWithoutSideEffects(imageElement, context);
    }

    private static bool HasUsableImageCandidateWithoutSideEffects(IEnumerable<string> candidates, ConversionContext context) {
        if (candidates == null) {
            return false;
        }

        foreach (string candidate in candidates) {
            if (CanUseImageCandidateWithoutSideEffects(candidate, context)) {
                return true;
            }
        }

        return false;
    }

    private static bool TryResolveUsableImageCandidate(IEnumerable<string> candidates, ConversionContext context, out string source) {
        source = string.Empty;
        if (candidates == null) {
            return false;
        }

        foreach (string candidate in candidates) {
            string resolved = candidate;
            if (string.IsNullOrWhiteSpace(resolved)) {
                continue;
            }

            if (!TryApplyBase64ImageHandling(ref resolved, context)) {
                continue;
            }

            if (string.IsNullOrWhiteSpace(resolved)) {
                continue;
            }

            source = resolved;
            return true;
        }

        return false;
    }

    private static bool CanUseImageCandidateWithoutSideEffects(string? source, ConversionContext context) {
        if (string.IsNullOrWhiteSpace(source)) {
            return false;
        }

        if (!HtmlImageDataUri.TryParse(source, out var dataUri) || !dataUri.IsBase64) {
            return true;
        }

        return context.Options.Base64Images switch {
            HtmlBase64ImageHandling.Include => true,
            HtmlBase64ImageHandling.Skip => false,
            HtmlBase64ImageHandling.SaveToFile => dataUri.TryDecodeBytes(out _),
            _ => throw new ArgumentOutOfRangeException(nameof(context.Options.Base64Images), context.Options.Base64Images, "Unknown base64 image handling mode.")
        };
    }

    private static ImageBlock CreateImageBlock(string src, IElement metadataElement, IElement? pictureElement = null, ConversionContext? context = null) {
        var image = new ImageBlock(
            src,
            GetAccessibleImageName(metadataElement),
            metadataElement.GetAttribute("title"));
        ApplyImageDimensions(metadataElement, image);
        if (pictureElement != null && context != null) {
            ApplyPictureMetadata(pictureElement, image, metadataElement, context);
        }
        return image;
    }

    private static bool TryCreateImageBlockFromNoscriptFallback(IElement element, ConversionContext context, out ImageBlock image) {
        image = null!;
        if (!TryResolveAssociatedNoscriptMediaElement(element, out var fallbackMediaElement)) {
            return false;
        }

        if (HasEffectiveTagName(fallbackMediaElement, context, "PICTURE")) {
            var pictureImage = ConvertPictureElement(fallbackMediaElement, context).OfType<ImageBlock>().FirstOrDefault();
            if (pictureImage == null) {
                return false;
            }

            image = MergeImageMetadata(element, pictureImage, FindFirstDescendantByEffectiveTagName(fallbackMediaElement, context, "IMG"));
            return true;
        }

        if (!TryResolveUsableImageCandidate(ResolveImageSourceCandidates(fallbackMediaElement, context), context, out string fallbackSrc)) {
            return false;
        }

        image = CreateMergedImageBlock(fallbackSrc, element, fallbackMediaElement);
        return true;
    }

    private static ImageBlock MergeImageMetadata(IElement preferredElement, ImageBlock fallbackImage, IElement? fallbackMetadataElement) {
        var merged = new ImageBlock(
            fallbackImage.Path,
            HasAccessibleNameDeclaration(preferredElement)
                ? GetAccessibleImageName(preferredElement)
                : fallbackImage.Alt,
            !string.IsNullOrWhiteSpace(preferredElement.GetAttribute("title")) ? preferredElement.GetAttribute("title") : fallbackImage.Title,
            fallbackImage.Width,
            fallbackImage.Height,
            fallbackImage.LinkUrl,
            fallbackImage.LinkTitle,
            fallbackImage.LinkTarget,
            fallbackImage.LinkRel) {
            Caption = fallbackImage.Caption,
            PictureFallbackPath = fallbackImage.PictureFallbackPath
        };
        CopyPictureSources(fallbackImage.PictureSources, merged.PictureSources);

        ApplyImageDimensions(preferredElement, merged);
        if ((merged.Width == null || merged.Height == null) && fallbackMetadataElement != null) {
            ApplyMissingImageDimensions(fallbackMetadataElement, merged);
        }

        return merged;
    }

    private static ImageBlock CreateMergedImageBlock(string src, IElement preferredMetadataElement, IElement fallbackMetadataElement) {
        var image = new ImageBlock(
            src,
            HasAccessibleNameDeclaration(preferredMetadataElement)
                ? GetAccessibleImageName(preferredMetadataElement)
                : GetAccessibleImageName(fallbackMetadataElement),
            !string.IsNullOrWhiteSpace(preferredMetadataElement.GetAttribute("title")) ? preferredMetadataElement.GetAttribute("title") : fallbackMetadataElement.GetAttribute("title"));
        ApplyImageDimensions(preferredMetadataElement, image);
        ApplyMissingImageDimensions(fallbackMetadataElement, image);
        return image;
    }

    private static bool HasAccessibleNameDeclaration(IElement element) =>
        element.HasAttribute("aria-labelledby")
        || element.HasAttribute("aria-label")
        || element.HasAttribute("alt")
        || element.HasAttribute("title")
        || element.TagName.Equals("SVG", StringComparison.OrdinalIgnoreCase)
           && element.Children.Any(static child => child.TagName.Equals("TITLE", StringComparison.OrdinalIgnoreCase));

    private static string? GetAccessibleImageName(IElement element) =>
        HasAccessibleNameDeclaration(element)
            ? HtmlAccessibilitySemantics.GetImageAccessibleName(element)
            : null;

    private static void ApplyPictureMetadata(IElement pictureElement, ImageBlock image, IElement? fallbackImageElement, ConversionContext context) {
        if (pictureElement == null || image == null || context == null) {
            return;
        }

        image.PictureSources.Clear();
        foreach (var source in CollectPictureSources(pictureElement, context)) {
            image.PictureSources.Add(source);
        }

        string fallbackPath = fallbackImageElement == null
            ? string.Empty
            : ResolvePictureFallbackImageSource(fallbackImageElement, context);
        if (!string.IsNullOrWhiteSpace(fallbackPath) && !TryApplyBase64ImageHandling(ref fallbackPath, context)) {
            fallbackPath = string.Empty;
        }

        image.PictureFallbackPath = string.IsNullOrWhiteSpace(fallbackPath) ? null : fallbackPath;
    }

    private static string ResolvePictureFallbackImageSource(IElement element, ConversionContext context) {
        string source = ResolveUrlAttributes(element, context, "src");
        string lazySource = ResolveUrlAttributes(element, context, "data-src", "data-original", "data-original-src", "data-lazy-src");

        if (IsLikelyPlaceholderImageSource(source) && !string.IsNullOrWhiteSpace(lazySource)) {
            return lazySource;
        }

        return !string.IsNullOrWhiteSpace(source) ? source : lazySource;
    }

    private static List<ImagePictureSource> CollectPictureSources(IElement pictureElement, ConversionContext context) {
        var sources = new List<ImagePictureSource>();
        foreach (var child in pictureElement.Children) {
            if (!HasEffectiveTagName(child, context, "SOURCE")) {
                continue;
            }

            string resolvedSrcSet = ResolveNormalizedSrcSetAttributes(child, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset");
            string resolved = ResolveUsableSrcSetCandidateAttributes(child, context, "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset");
            if (string.IsNullOrWhiteSpace(resolved)) {
                resolved = ResolveUrlAttributes(child, context, "src", "data-src", "data-original-src", "data-lazy-src");
            }

            if (string.IsNullOrWhiteSpace(resolved)) {
                continue;
            }

            if (!TryApplyBase64ImageHandling(ref resolved, context)) {
                continue;
            }

            resolvedSrcSet = ApplyBase64ImageHandlingToSrcSet(resolvedSrcSet, context);
            sources.Add(new ImagePictureSource(
                resolved,
                child.GetAttribute("media"),
                child.GetAttribute("type"),
                child.GetAttribute("sizes"),
                resolvedSrcSet));
        }

        return sources;
    }

    private static string ResolveUsableSrcSetCandidateAttributes(IElement element, ConversionContext context, params string[] attributeNames) {
        if (element == null || attributeNames == null || attributeNames.Length == 0) {
            return string.Empty;
        }

        for (int i = 0; i < attributeNames.Length; i++) {
            IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlImageSourceResolver.ResolveSrcSetCandidates(
                element.GetAttribute(attributeNames[i]),
                context.Options.BaseUri,
                context.Options.ResourceUrlPolicy ?? context.Options.UrlPolicy);

            foreach (HtmlSrcSetCandidate candidate in candidates) {
                string resolved = candidate.Url;
                if (!TryApplyBase64ImageHandling(ref resolved, context)) {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(resolved)) {
                    return resolved;
                }
            }
        }

        return string.Empty;
    }

    private static string ApplyBase64ImageHandlingToSrcSet(string? srcSet, ConversionContext context) {
        if (string.IsNullOrWhiteSpace(srcSet)) {
            return string.Empty;
        }

        var parts = new List<string>();
        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(srcSet)) {
            string originalSource = candidate.Url;
            string source = candidate.Url;
            if (!TryApplyBase64ImageHandling(ref source, context) || string.IsNullOrWhiteSpace(source)) {
                continue;
            }

            source = EncodeSavedSrcSetSource(originalSource, source, context);
            parts.Add(string.IsNullOrWhiteSpace(candidate.Descriptor)
                ? source
                : source + " " + candidate.Descriptor);
        }

        return string.Join(", ", parts);
    }

    private static string EncodeSavedSrcSetSource(string originalSource, string source, ConversionContext context) {
        if (context.Options.Base64Images != HtmlBase64ImageHandling.SaveToFile
            || !IsBase64ImageDataUri(originalSource)
            || string.Equals(originalSource, source, StringComparison.Ordinal)) {
            return source;
        }

        return EncodeMarkdownUrlWhitespace(source);
    }

    private static string EncodeMarkdownUrlWhitespace(string source) {
        StringBuilder? builder = null;
        for (int i = 0; i < source.Length; i++) {
            char value = source[i];
            if (!char.IsWhiteSpace(value)) {
                builder?.Append(value);
                continue;
            }

            builder ??= new StringBuilder(source.Length + 8);
            if (builder.Length == 0 && i > 0) {
                builder.Append(source, 0, i);
            }

            if (value == ' ') {
                builder.Append("%20");
            } else {
                builder.Append(Uri.HexEscape(value));
            }
        }

        return builder?.ToString() ?? source;
    }

    private static void CopyPictureSources(IEnumerable<ImagePictureSource> sourceItems, IList<ImagePictureSource> targetItems) {
        if (sourceItems == null || targetItems == null) {
            return;
        }

        targetItems.Clear();
        foreach (var source in sourceItems) {
            if (source == null || string.IsNullOrWhiteSpace(source.Path)) {
                continue;
            }

            targetItems.Add(new ImagePictureSource(source.Path, source.Media, source.Type, source.Sizes, source.SrcSet));
        }
    }

    private static void ApplyMissingImageDimensions(IElement element, ImageBlock image) {
        if (image.Width == null
            && (TryParseImageDimension(element.GetAttribute("width"), out double width)
                || TryParseStyleDimension(element.GetAttribute("style"), "width", out width))) {
            image.Width = width;
        }

        if (image.Height == null
            && (TryParseImageDimension(element.GetAttribute("height"), out double height)
                || TryParseStyleDimension(element.GetAttribute("style"), "height", out height))) {
            image.Height = height;
        }
    }

    private static bool TryResolveAssociatedNoscriptMediaElement(IElement element, out IElement mediaElement) {
        mediaElement = null!;
        foreach (var noscriptElement in EnumerateAssociatedNoscriptElements(element)) {
            if (TryResolveNoscriptMediaElement(noscriptElement, out mediaElement)) {
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<IElement> EnumerateAssociatedNoscriptElements(IElement element) {
        var visited = new HashSet<IElement>();
        IElement? current = element.ParentElement;
        int depth = 0;
        while (current != null && depth < 3) {
            foreach (var child in current.Children) {
                if (child.TagName.Equals("NOSCRIPT", StringComparison.OrdinalIgnoreCase) && visited.Add(child)) {
                    yield return child;
                }
            }

            if (!IsPotentialMediaFallbackContainer(current)) {
                yield break;
            }

            current = current.ParentElement;
            depth++;
        }
    }

    private static bool IsPotentialMediaFallbackContainer(IElement element) {
        if (element == null) {
            return false;
        }

        return element.TagName.Equals("A", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("DIV", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("SPAN", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)
               || element.TagName.Equals("FIGURE", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryResolveNoscriptMediaElement(IElement noscriptElement, out IElement mediaElement) {
        mediaElement = null!;
        foreach (string html in EnumerateNoscriptHtmlCandidates(noscriptElement)) {
            var document = HtmlDocumentParser.ParseDocument($"<body>{html}</body>");
            IElement? parsedMediaElement = document.QuerySelector("picture") ?? document.QuerySelector("img");
            if (parsedMediaElement != null) {
                mediaElement = parsedMediaElement;
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<string> EnumerateNoscriptHtmlCandidates(IElement noscriptElement) {
        if (noscriptElement == null) {
            yield break;
        }

        string innerHtml = noscriptElement.InnerHtml;
        if (!string.IsNullOrWhiteSpace(innerHtml)) {
            yield return innerHtml;
        }

        string textContent = noscriptElement.TextContent;
        if (!string.IsNullOrWhiteSpace(textContent) && !string.Equals(textContent, innerHtml, StringComparison.Ordinal)) {
            yield return textContent;
        }
    }

    private static bool IsLikelyPlaceholderImageSource(string? source) {
        if (string.IsNullOrWhiteSpace(source)) {
            return false;
        }

        string value = source!.Trim();
        return value.Equals("about:blank", StringComparison.OrdinalIgnoreCase)
               || value.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase)
               || value.Contains("transparent.gif", StringComparison.OrdinalIgnoreCase)
               || value.Contains("spacer.gif", StringComparison.OrdinalIgnoreCase)
               || value.Contains("blank.gif", StringComparison.OrdinalIgnoreCase)
               || value.Contains("pixel.gif", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryCreateLinkedImageBlockFromAnchor(IElement anchorElement, ConversionContext context, out ImageBlock image) {
        image = null!;
        if (anchorElement == null || !HasEffectiveTagName(anchorElement, context, "A")) {
            return false;
        }

        string href = ResolveUrl(anchorElement.GetAttribute("href"), context);
        if (string.IsNullOrWhiteSpace(href)) {
            return false;
        }

        if (!TryResolveAnchorMediaElement(anchorElement, context, out var mediaElement)) {
            return false;
        }

        if (HasEffectiveTagName(mediaElement, context, "IMG")
            && TryCreateImageBlock(mediaElement, context, out image)) {
            image.LinkUrl = href;
            image.LinkTitle = anchorElement.GetAttribute("title");
            image.LinkTarget = anchorElement.GetAttribute("target");
            image.LinkRel = anchorElement.GetAttribute("rel");
            return true;
        }

        if (!HasEffectiveTagName(mediaElement, context, "PICTURE")) {
            return false;
        }

        var pictureImage = ConvertPictureElement(mediaElement, context).OfType<ImageBlock>().FirstOrDefault();
        if (pictureImage == null) {
            return false;
        }

        pictureImage.LinkUrl = href;
        pictureImage.LinkTitle = anchorElement.GetAttribute("title");
        pictureImage.LinkTarget = anchorElement.GetAttribute("target");
        pictureImage.LinkRel = anchorElement.GetAttribute("rel");
        image = pictureImage;
        return true;
    }

    private static bool TryParseStyleDimension(string? style, string propertyName, out double value) {
        value = default;
        return TryParseImageDimension(TryGetStyleDeclarationValue(style, propertyName), out value);
    }

    private static string? TryGetStyleDeclarationValue(string? style, string propertyName) {
        if (string.IsNullOrWhiteSpace(style) || string.IsNullOrWhiteSpace(propertyName)) {
            return null;
        }

        foreach (var declaration in style!.Split(';')) {
            if (string.IsNullOrWhiteSpace(declaration)) {
                continue;
            }

            int separatorIndex = declaration.IndexOf(':');
            if (separatorIndex <= 0) {
                continue;
            }

            string name = declaration.Substring(0, separatorIndex).Trim();
            if (!name.Equals(propertyName, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            string value = declaration.Substring(separatorIndex + 1).Trim();
            return value.Length == 0 ? null : value;
        }

        return null;
    }

    private static bool TryParseImageDimension(string? rawValue, out double value) {
        value = default;
        if (string.IsNullOrWhiteSpace(rawValue)) {
            return false;
        }

        string normalized = rawValue!.Trim();
        if (normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase)) {
            normalized = normalized.Substring(0, normalized.Length - 2).Trim();
        }

        if (normalized.Length == 0
            || normalized.IndexOf('%') >= 0
            || normalized.IndexOf("calc(", StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf("var(", StringComparison.OrdinalIgnoreCase) >= 0) {
            return false;
        }

        return double.TryParse(
            normalized,
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture,
            out value);
    }
}

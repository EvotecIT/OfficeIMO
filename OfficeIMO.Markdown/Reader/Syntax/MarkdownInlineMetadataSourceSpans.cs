using System.Runtime.CompilerServices;

namespace OfficeIMO.Markdown;

internal static class MarkdownInlineMetadataSourceSpans {
    private sealed class LinkHolder {
        public MarkdownSourceSpan? TargetSpan;
        public MarkdownSourceSpan? TitleSpan;
        public MarkdownSourceSpan? HtmlTargetSpan;
        public MarkdownSourceSpan? HtmlRelSpan;
    }

    private class ImageHolder {
        public MarkdownSourceSpan? AltSpan;
        public MarkdownSourceSpan? SourceSpan;
        public MarkdownSourceSpan? TitleSpan;
    }

    private sealed class ImageLinkHolder : ImageHolder {
        public MarkdownSourceSpan? LinkTargetSpan;
        public MarkdownSourceSpan? LinkTitleSpan;
    }

    private static readonly ConditionalWeakTable<LinkInline, LinkHolder> _linkSpans = new();
    private static readonly ConditionalWeakTable<ImageInline, ImageHolder> _imageSpans = new();
    private static readonly ConditionalWeakTable<ImageLinkInline, ImageLinkHolder> _imageLinkSpans = new();

    internal static void SetLinkParts(
        LinkInline? inline,
        MarkdownSourceSpan? targetSpan,
        MarkdownSourceSpan? titleSpan,
        MarkdownSourceSpan? htmlTargetSpan = null,
        MarkdownSourceSpan? htmlRelSpan = null) {
        if (inline == null) {
            return;
        }

        if (!targetSpan.HasValue && !titleSpan.HasValue && !htmlTargetSpan.HasValue && !htmlRelSpan.HasValue) {
            return;
        }

        _linkSpans.Remove(inline);
        _linkSpans.Add(inline, new LinkHolder {
            TargetSpan = targetSpan,
            TitleSpan = titleSpan,
            HtmlTargetSpan = htmlTargetSpan,
            HtmlRelSpan = htmlRelSpan
        });
    }

    internal static MarkdownSourceSpan? GetLinkTargetSpan(LinkInline? inline) =>
        inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.TargetSpan : null;

    internal static MarkdownSourceSpan? GetLinkTitleSpan(LinkInline? inline) =>
        inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.TitleSpan : null;

    internal static MarkdownSourceSpan? GetLinkHtmlTargetSpan(LinkInline? inline) =>
        inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.HtmlTargetSpan : null;

    internal static MarkdownSourceSpan? GetLinkHtmlRelSpan(LinkInline? inline) =>
        inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.HtmlRelSpan : null;

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

        _imageSpans.Remove(inline);
        _imageSpans.Add(inline, new ImageHolder {
            AltSpan = altSpan,
            SourceSpan = sourceSpan,
            TitleSpan = titleSpan
        });
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

        _imageLinkSpans.Remove(inline);
        _imageLinkSpans.Add(inline, new ImageLinkHolder {
            AltSpan = altSpan,
            SourceSpan = sourceSpan,
            TitleSpan = imageTitleSpan,
            LinkTargetSpan = linkTargetSpan,
            LinkTitleSpan = linkTitleSpan
        });
    }

    internal static MarkdownSourceSpan? GetImageAltSpan(ImageInline? inline) =>
        inline != null && _imageSpans.TryGetValue(inline, out var holder) ? holder.AltSpan : null;

    internal static MarkdownSourceSpan? GetImageSourceSpan(ImageInline? inline) =>
        inline != null && _imageSpans.TryGetValue(inline, out var holder) ? holder.SourceSpan : null;

    internal static MarkdownSourceSpan? GetImageTitleSpan(ImageInline? inline) =>
        inline != null && _imageSpans.TryGetValue(inline, out var holder) ? holder.TitleSpan : null;

    internal static MarkdownSourceSpan? GetImageAltSpan(ImageLinkInline? inline) =>
        inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.AltSpan : null;

    internal static MarkdownSourceSpan? GetImageSourceSpan(ImageLinkInline? inline) =>
        inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.SourceSpan : null;

    internal static MarkdownSourceSpan? GetImageTitleSpan(ImageLinkInline? inline) =>
        inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.TitleSpan : null;

    internal static MarkdownSourceSpan? GetImageLinkTargetSpan(ImageLinkInline? inline) =>
        inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.LinkTargetSpan : null;

    internal static MarkdownSourceSpan? GetImageLinkTitleSpan(ImageLinkInline? inline) =>
        inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.LinkTitleSpan : null;
}

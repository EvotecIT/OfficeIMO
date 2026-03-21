using System.Runtime.CompilerServices;

namespace OfficeIMO.Markdown;

internal static class MarkdownInlineMetadataSourceSpans {
    private sealed class LinkState {
        public MarkdownSourceSpan? TargetSpan;
        public MarkdownSourceSpan? TitleSpan;
        public MarkdownSourceSpan? HtmlTargetSpan;
        public MarkdownSourceSpan? HtmlRelSpan;
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

    // These tables hold weak references to markdown inline keys, so entries disappear when the
    // owning inline objects are no longer referenced by the parse result or callers.
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

        var holder = _linkSpans.GetValue(inline, static _ => new LinkHolder());
        holder.State = new LinkState {
            TargetSpan = targetSpan,
            TitleSpan = titleSpan,
            HtmlTargetSpan = htmlTargetSpan,
            HtmlRelSpan = htmlRelSpan
        };
    }

    internal static MarkdownSourceSpan? GetLinkTargetSpan(LinkInline? inline) =>
        inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.State?.TargetSpan : null;

    internal static MarkdownSourceSpan? GetLinkTitleSpan(LinkInline? inline) =>
        inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.State?.TitleSpan : null;

    internal static MarkdownSourceSpan? GetLinkHtmlTargetSpan(LinkInline? inline) =>
        inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.State?.HtmlTargetSpan : null;

    internal static MarkdownSourceSpan? GetLinkHtmlRelSpan(LinkInline? inline) =>
        inline != null && _linkSpans.TryGetValue(inline, out var holder) ? holder.State?.HtmlRelSpan : null;

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
    }

    internal static MarkdownSourceSpan? GetImageAltSpan(ImageInline? inline) =>
        inline != null && _imageSpans.TryGetValue(inline, out var holder) ? holder.State?.AltSpan : null;

    internal static MarkdownSourceSpan? GetImageSourceSpan(ImageInline? inline) =>
        inline != null && _imageSpans.TryGetValue(inline, out var holder) ? holder.State?.SourceSpan : null;

    internal static MarkdownSourceSpan? GetImageTitleSpan(ImageInline? inline) =>
        inline != null && _imageSpans.TryGetValue(inline, out var holder) ? holder.State?.TitleSpan : null;

    internal static MarkdownSourceSpan? GetImageAltSpan(ImageLinkInline? inline) =>
        inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.State?.AltSpan : null;

    internal static MarkdownSourceSpan? GetImageSourceSpan(ImageLinkInline? inline) =>
        inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.State?.SourceSpan : null;

    internal static MarkdownSourceSpan? GetImageTitleSpan(ImageLinkInline? inline) =>
        inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.State?.TitleSpan : null;

    internal static MarkdownSourceSpan? GetImageLinkTargetSpan(ImageLinkInline? inline) =>
        inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.State?.LinkTargetSpan : null;

    internal static MarkdownSourceSpan? GetImageLinkTitleSpan(ImageLinkInline? inline) =>
        inline != null && _imageLinkSpans.TryGetValue(inline, out var holder) ? holder.State?.LinkTitleSpan : null;
}

using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private double ResolveReplacedImageBoxWidth(IElement element, HtmlRenderBoxStyle style) {
        TryResolveImageSource(
            element.GetAttribute("src"),
            HtmlRenderStyleResolver.DescribeSource(element),
            out _,
            out _,
            out OfficeImageInfo? imageInfo,
            reportDiagnostics: false);
        bool hasIntrinsicSize = imageInfo != null && imageInfo.Width > 0 && imageInfo.Height > 0;
        double intrinsicWidth = hasIntrinsicSize
            ? imageInfo!.Width * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiX)
            : 300D;
        double intrinsicHeight = hasIntrinsicSize
            ? imageInfo!.Height * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiY)
            : 150D;
        ReplacedContentSize size = ResolveReplacedContentSize(style, intrinsicWidth, intrinsicHeight, hasIntrinsicSize);
        double boxWidth = size.Width + style.HorizontalInsets;
        double boxHeight = size.Height + style.VerticalInsets;
        EnsureReplacedBoxSize(boxWidth, boxHeight);
        return boxWidth;
    }

    private ReplacedContentSize ResolveReplacedContentSize(
        HtmlRenderBoxStyle style,
        double intrinsicWidth,
        double intrinsicHeight,
        bool hasIntrinsicSize) {
        double intrinsicRatio = intrinsicWidth / intrinsicHeight;
        bool useAuthoredRatio = style.AspectRatio.HasValue && (!style.AspectRatioPrefersIntrinsic || !hasIntrinsicSize);
        double preferredRatio = useAuthoredRatio ? style.AspectRatio!.Value : intrinsicRatio;
        double? declaredWidth = ResolveReplacedDimension(style.ExplicitWidth, style.PaddingLeft + style.PaddingRight + style.BorderLeftWidth + style.BorderRightWidth, style.BorderBox);
        double? declaredHeight = ResolveReplacedDimension(style.ExplicitHeight, style.PaddingTop + style.PaddingBottom + style.BorderTopWidth + style.BorderBottomWidth, style.BorderBox);

        double width;
        double height;
        if (declaredWidth.HasValue && declaredHeight.HasValue) {
            width = declaredWidth.Value;
            height = declaredHeight.Value;
        } else if (declaredWidth.HasValue) {
            width = declaredWidth.Value;
            height = width / preferredRatio;
        } else if (declaredHeight.HasValue) {
            height = declaredHeight.Value;
            width = height * preferredRatio;
        } else {
            width = intrinsicWidth;
            height = useAuthoredRatio ? width / preferredRatio : intrinsicHeight;
        }

        double minimumWidth = ResolveReplacedDimension(style.MinWidth, style.HorizontalInsets, style.BorderBox) ?? 0.01D;
        double maximumWidth = ResolveReplacedDimension(style.MaxWidth, style.HorizontalInsets, style.BorderBox) ?? double.PositiveInfinity;
        double minimumHeight = ResolveReplacedDimension(style.MinHeight, style.VerticalInsets, style.BorderBox) ?? 0.01D;
        double maximumHeight = ResolveReplacedDimension(style.MaxHeight, style.VerticalInsets, style.BorderBox) ?? double.PositiveInfinity;
        if (declaredWidth.HasValue && declaredHeight.HasValue) {
            width = ClampWithMinimumPrecedence(width, minimumWidth, maximumWidth);
            height = ClampWithMinimumPrecedence(height, minimumHeight, maximumHeight);
        } else {
            ConstrainProportionalSize(ref width, ref height, minimumWidth, maximumWidth, minimumHeight, maximumHeight);
        }
        return new ReplacedContentSize(Math.Max(0.01D, width), Math.Max(0.01D, height));
    }

    private ReplacedObjectPlacement ResolveReplacedObjectPlacement(
        HtmlRenderBoxStyle style,
        double contentWidth,
        double contentHeight,
        double intrinsicWidth,
        double intrinsicHeight) {
        double objectWidth = contentWidth;
        double objectHeight = contentHeight;
        if (style.ObjectFit != "fill") {
            double containScale = Math.Min(contentWidth / intrinsicWidth, contentHeight / intrinsicHeight);
            double scale = style.ObjectFit == "cover"
                ? Math.Max(contentWidth / intrinsicWidth, contentHeight / intrinsicHeight)
                : style.ObjectFit == "none" || style.ObjectFit == "scale-down" && containScale >= 1D
                    ? 1D
                    : containScale;
            objectWidth = intrinsicWidth * scale;
            objectHeight = intrinsicHeight * scale;
        }

        if (!HtmlCssReplacedElementParser.TryResolveObjectPosition(
                style.ObjectPosition,
                contentWidth,
                contentHeight,
                objectWidth,
                objectHeight,
                style.Font.Size,
                _options.DefaultFontSize,
                out double objectX,
                out double objectY)) {
            objectX = (contentWidth - objectWidth) / 2D;
            objectY = (contentHeight - objectHeight) / 2D;
        }

        double visibleLeft = Math.Max(0D, objectX);
        double visibleTop = Math.Max(0D, objectY);
        double visibleRight = Math.Min(contentWidth, objectX + objectWidth);
        double visibleBottom = Math.Min(contentHeight, objectY + objectHeight);
        if (visibleRight <= visibleLeft + 0.000001D || visibleBottom <= visibleTop + 0.000001D) return default;

        double cropLeft = Math.Max(0D, (visibleLeft - objectX) / objectWidth);
        double cropTop = Math.Max(0D, (visibleTop - objectY) / objectHeight);
        double cropRight = Math.Max(0D, (objectX + objectWidth - visibleRight) / objectWidth);
        double cropBottom = Math.Max(0D, (objectY + objectHeight - visibleBottom) / objectHeight);
        PreserveVisibleCrop(ref cropLeft, ref cropRight);
        PreserveVisibleCrop(ref cropTop, ref cropBottom);
        OfficeImageSourceCrop crop = cropLeft > 0D || cropTop > 0D || cropRight > 0D || cropBottom > 0D
            ? OfficeImageSourceCrop.FromStrictFractions(cropLeft, cropTop, cropRight, cropBottom)
            : default;
        return new ReplacedObjectPlacement(
            visibleLeft,
            visibleTop,
            visibleRight - visibleLeft,
            visibleBottom - visibleTop,
            crop);
    }

    private void ReportReplacedElementFallbacks(HtmlRenderBoxStyle style, IElement element) {
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        if (style.UnsupportedReplacedElementLayout.Length > 0 && _reportedReplacedElementFallbacks.Add(source)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.ReplacedElementValueUnsupported,
                "A replaced-element sizing or object-placement value used a deterministic fallback.",
                HtmlDiagnosticSeverity.Warning,
                source,
                style.UnsupportedReplacedElementLayout);
        }
    }

    private static double? ResolveReplacedDimension(double? value, double insets, bool borderBox) {
        if (!value.HasValue) return null;
        return Math.Max(0.01D, borderBox ? value.Value - insets : value.Value);
    }

    private static double ClampWithMinimumPrecedence(double value, double minimum, double maximum) =>
        Math.Max(minimum, Math.Min(value, maximum));

    private void EnsureReplacedBoxSize(double width, double height) {
        double pixelWidth = Math.Ceiling(width * _options.Scale);
        double pixelHeight = Math.Ceiling(height * _options.Scale);
        if (double.IsNaN(pixelWidth) || double.IsInfinity(pixelWidth)
            || double.IsNaN(pixelHeight) || double.IsInfinity(pixelHeight)
            || pixelWidth > _options.MaxSurfaceWidth
            || pixelHeight > _options.MaxSurfaceHeight) {
            throw new InvalidOperationException("HTML replaced content exceeded the configured maximum image surface dimensions.");
        }
    }

    private static void PreserveVisibleCrop(ref double start, ref double end) {
        double total = start + end;
        double maximum = 1D - OfficeImageSourceCrop.MinimumVisibleRatio;
        if (total <= maximum) return;
        double scale = maximum / total;
        start *= scale;
        end *= scale;
    }

    private static void ConstrainProportionalSize(
        ref double width,
        ref double height,
        double minimumWidth,
        double maximumWidth,
        double minimumHeight,
        double maximumHeight) {
        width = Math.Max(0.01D, width);
        height = Math.Max(0.01D, height);
        double minimumScale = Math.Max(minimumWidth / width, minimumHeight / height);
        double maximumScale = Math.Min(maximumWidth / width, maximumHeight / height);
        double scale = Math.Max(1D, minimumScale);
        if (maximumScale >= minimumScale) scale = Math.Min(scale, maximumScale);
        width *= scale;
        height *= scale;
    }

    private readonly struct ReplacedContentSize {
        internal ReplacedContentSize(double width, double height) {
            Width = width;
            Height = height;
        }

        internal double Width { get; }
        internal double Height { get; }
    }

    private readonly struct ReplacedObjectPlacement {
        internal ReplacedObjectPlacement(double x, double y, double width, double height, OfficeImageSourceCrop sourceCrop) {
            X = x;
            Y = y;
            Width = width;
            Height = height;
            SourceCrop = sourceCrop;
        }

        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal OfficeImageSourceCrop SourceCrop { get; }
        internal bool IsVisible => Width > 0D && Height > 0D;
    }
}

using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock LayoutImage(IElement element, double containingWidth, HtmlRenderBoxStyle style) {
        string sourceDescription = HtmlRenderStyleResolver.DescribeSource(element);
        string? source = element.GetAttribute("src");
        TryResolveImageSource(source, sourceDescription, out byte[]? bytes, out string contentType, out OfficeImageInfo? imageInfo);
        bool hasIntrinsicSize = imageInfo != null && imageInfo.Width > 0 && imageInfo.Height > 0;
        double intrinsicWidth = hasIntrinsicSize
            ? imageInfo!.Width * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiX)
            : 300D;
        double intrinsicHeight = hasIntrinsicSize
            ? imageInfo!.Height * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiY)
            : 150D;
        ReplacedContentSize contentSize = ResolveReplacedContentSize(style, intrinsicWidth, intrinsicHeight, hasIntrinsicSize);
        double boxWidth = contentSize.Width + style.HorizontalInsets;
        double boxHeight = contentSize.Height + style.VerticalInsets;
        var visuals = new List<HtmlRenderVisual>();
        var objectVisuals = new List<HtmlRenderVisual>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        double imageX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double imageY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        string? link = element.ParentElement != null && string.Equals(element.ParentElement.TagName, "a", StringComparison.OrdinalIgnoreCase)
            ? ResolveSafeLink(element.ParentElement.GetAttribute("href"), element.ParentElement)
            : null;
        string? alternativeText = element.GetAttribute("alt");
        ReplacedObjectPlacement placement = ResolveReplacedObjectPlacement(
            style,
            contentSize.Width,
            contentSize.Height,
            intrinsicWidth,
            intrinsicHeight);
        if (bytes != null && bytes.Length > 0 && placement.IsVisible) {
            objectVisuals.Add(new HtmlRenderImage(
                bytes,
                contentType,
                imageX + placement.X,
                imageY + placement.Y,
                placement.Width,
                placement.Height,
                objectVisuals.Count,
                alternativeText,
                link,
                sourceDescription,
                placement.SourceCrop));
            if (!OfficeRasterImageDecoder.TryDecode(bytes, out _) && !string.Equals(contentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.RasterDecoderUnavailable, "The image can be retained for SVG/PDF but the dependency-free PNG backend cannot decode this image format yet.", HtmlDiagnosticSeverity.Warning, sourceDescription, contentType);
            }
        } else if (placement.IsVisible) {
            OfficeShape placeholder = OfficeShape.Rectangle(placement.Width, placement.Height);
            placeholder.FillColor = OfficeColor.FromRgb(245, 245, 245);
            placeholder.StrokeColor = OfficeColor.FromRgb(160, 160, 160);
            placeholder.StrokeWidth = 1D;
            objectVisuals.Add(new HtmlRenderShape(placeholder, imageX + placement.X, imageY + placement.Y, objectVisuals.Count, link, sourceDescription));
            if (!string.IsNullOrWhiteSpace(alternativeText)) {
                double textHeight = Math.Min(placement.Height, style.LineHeight);
                objectVisuals.Add(new HtmlRenderText(alternativeText!, imageX + placement.X + 4D, imageY + placement.Y + 4D, Math.Max(1D, placement.Width - 8D), Math.Max(1D, textHeight), style.Font, style.Color, OfficeTextAlignment.Left, style.LineHeight, objectVisuals.Count, link, sourceDescription, "figure-alternative-text"));
            }
        }
        HtmlResolvedBorderRadii outerRadii = ResolveBoxRadii(style, boxWidth, boxHeight, element, sourceDescription);
        HtmlResolvedBorderRadii contentRadii = outerRadii.Inset(
            style.BorderLeftWidth + style.PaddingLeft,
            style.BorderTopWidth + style.PaddingTop,
            style.BorderRightWidth + style.PaddingRight,
            style.BorderBottomWidth + style.PaddingBottom,
            contentSize.Width,
            contentSize.Height);
        AddBoxClipVisuals(
            visuals,
            objectVisuals,
            imageX,
            imageY,
            contentSize.Width,
            contentSize.Height,
            contentRadii,
            sourceDescription + ":content-clip");
        ReportReplacedElementFallbacks(style, element);
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);

        double outerHeight = style.MarginTop + boxHeight + style.MarginBottom;
        return new HtmlRenderFlowBlock(containingWidth, outerHeight, visuals, style.BreakBefore, style.BreakAfter, style.AvoidBreakInside, sourceDescription, pageName: style.PageName);
    }

    private double ResolveFloatingImageOuterWidth(IElement element, HtmlRenderBoxStyle style) {
        return Math.Max(1D, style.MarginLeft + ResolveReplacedImageBoxWidth(element, style) + style.MarginRight);
    }

    private bool TryResolveImageSource(
        string? source,
        string sourceDescription,
        out byte[]? bytes,
        out string contentType,
        out OfficeImageInfo? imageInfo,
        bool reportDiagnostics = true) {
        bytes = null;
        contentType = string.Empty;
        imageInfo = null;
        string resolvedSource = HtmlUrlPolicyEvaluator.ResolveUrl(source, _baseUri, _resourceUrlPolicy);
        string extension = string.Empty;
        if (_resources.TryGet(source, resolvedSource, out HtmlResolvedResource resolvedResource)) {
            bytes = resolvedResource.Bytes;
            contentType = resolvedResource.ContentType;
            extension = OfficeImageInfo.GetDefaultExtension(OfficeImageInfo.FromMimeType(contentType));
        } else if (HtmlImageDataUri.TryParse(source, out HtmlImageDataUri dataUri) && dataUri.TryDecodeBytes(out byte[] decoded)) {
            bytes = decoded;
            contentType = dataUri.MediaType;
            extension = dataUri.FileExtension;
        } else if (reportDiagnostics && !string.IsNullOrWhiteSpace(source) && !_resources.WasAttempted(source, resolvedSource)) {
            string code = resolvedSource.Length == 0 ? "ImageResourceRejectedByPolicy" : HtmlRenderDiagnosticCodes.ExternalImagePending;
            string message = resolvedSource.Length == 0
                ? "An image was rejected before entering the rendered document."
                : "Synchronous rendering does not load external images; use RenderAsync with an application-supplied resolver or provide a data URI.";
            _diagnostics.Add(ComponentName, code, message, HtmlDiagnosticSeverity.Warning, sourceDescription, source);
        }

        if (bytes == null || bytes.Length == 0) {
            return false;
        }

        if (OfficeImageReader.TryIdentify(bytes, extension, out OfficeImageInfo identified)) {
            imageInfo = identified;
        }

        return true;
    }
}

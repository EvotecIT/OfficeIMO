using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock LayoutImage(IElement element, double containingWidth, HtmlRenderBoxStyle style) {
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        string sourceDescription = HtmlRenderStyleResolver.DescribeSource(element);
        string? source = element.GetAttribute("src");
        byte[]? bytes = null;
        string contentType = string.Empty;
        OfficeImageInfo? imageInfo = null;
        string resolvedSource = HtmlUrlPolicyEvaluator.ResolveUrl(source, _baseUri, _options.UrlPolicy);
        if (_resources.TryGet(source, resolvedSource, out HtmlResolvedResource resolvedResource)) {
            bytes = resolvedResource.Bytes;
            contentType = resolvedResource.ContentType;
            if (OfficeImageReader.TryIdentify(bytes, OfficeImageInfo.GetDefaultExtension(OfficeImageInfo.FromMimeType(contentType)), out OfficeImageInfo identified)) imageInfo = identified;
        } else if (HtmlImageDataUri.TryParse(source, out HtmlImageDataUri dataUri) && dataUri.TryDecodeBytes(out byte[] decoded)) {
            bytes = decoded;
            contentType = dataUri.MediaType;
            if (OfficeImageReader.TryIdentify(bytes, dataUri.FileExtension, out OfficeImageInfo identified)) imageInfo = identified;
        } else if (!string.IsNullOrWhiteSpace(source) && !_resources.WasAttempted(source, resolvedSource)) {
            string code = resolvedSource.Length == 0 ? "ImageResourceRejectedByPolicy" : HtmlRenderDiagnosticCodes.ExternalImagePending;
            string message = resolvedSource.Length == 0
                ? "An image was rejected before entering the rendered document."
                : "Synchronous rendering does not load external images; use RenderAsync with an application-supplied resolver or provide a data URI.";
            _diagnostics.Add(ComponentName, code, message, HtmlDiagnosticSeverity.Warning, sourceDescription, source);
        }

        double intrinsicWidth = imageInfo != null && imageInfo.Width > 0 ? imageInfo.Width * HtmlRenderOptions.CssPixelsPerInch / imageInfo.DpiX : 300D;
        double intrinsicHeight = imageInfo != null && imageInfo.Height > 0 ? imageInfo.Height * HtmlRenderOptions.CssPixelsPerInch / imageInfo.DpiY : 150D;
        double imageWidth = style.ExplicitWidth ?? intrinsicWidth;
        double imageHeight = style.ExplicitHeight ?? (style.ExplicitWidth.HasValue && intrinsicWidth > 0D ? imageWidth * intrinsicHeight / intrinsicWidth : intrinsicHeight);
        if (!style.ExplicitWidth.HasValue && style.ExplicitHeight.HasValue && intrinsicHeight > 0D) imageWidth = imageHeight * intrinsicWidth / intrinsicHeight;
        double maximumContentWidth = Math.Max(1D, availableWidth - style.HorizontalInsets);
        if (imageWidth > maximumContentWidth) {
            double scale = maximumContentWidth / imageWidth;
            imageWidth = maximumContentWidth;
            imageHeight *= scale;
        }

        imageWidth = Math.Max(1D, imageWidth);
        imageHeight = Math.Max(1D, imageHeight);
        double boxWidth = Math.Min(availableWidth, imageWidth + style.HorizontalInsets);
        double boxHeight = imageHeight + style.VerticalInsets;
        var visuals = new List<HtmlRenderVisual>();
        AddBoxShape(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        double imageX = style.MarginLeft + style.BorderWidth + style.PaddingLeft;
        double imageY = style.MarginTop + style.BorderWidth + style.PaddingTop;
        string? link = element.ParentElement != null && string.Equals(element.ParentElement.TagName, "a", StringComparison.OrdinalIgnoreCase)
            ? ResolveSafeLink(element.ParentElement.GetAttribute("href"), element.ParentElement)
            : null;
        string? alternativeText = element.GetAttribute("alt");
        if (bytes != null && bytes.Length > 0) {
            visuals.Add(new HtmlRenderImage(bytes, contentType, imageX, imageY, imageWidth, imageHeight, visuals.Count, alternativeText, link, sourceDescription));
            if (!OfficeRasterImageDecoder.TryDecode(bytes, out _) && !string.Equals(contentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.RasterDecoderUnavailable, "The image can be retained for SVG/PDF but the dependency-free PNG backend cannot decode this image format yet.", HtmlDiagnosticSeverity.Warning, sourceDescription, contentType);
            }
        } else {
            OfficeShape placeholder = OfficeShape.Rectangle(imageWidth, imageHeight);
            placeholder.FillColor = OfficeColor.FromRgb(245, 245, 245);
            placeholder.StrokeColor = OfficeColor.FromRgb(160, 160, 160);
            placeholder.StrokeWidth = 1D;
            visuals.Add(new HtmlRenderShape(placeholder, imageX, imageY, visuals.Count, link, sourceDescription));
            if (!string.IsNullOrWhiteSpace(alternativeText)) {
                double textHeight = Math.Min(imageHeight, style.LineHeight);
                visuals.Add(new HtmlRenderText(alternativeText!, imageX + 4D, imageY + 4D, Math.Max(1D, imageWidth - 8D), Math.Max(1D, textHeight), style.Font, style.Color, OfficeTextAlignment.Left, style.LineHeight, visuals.Count, link, sourceDescription, "figure-alternative-text"));
            }
        }

        double outerHeight = style.MarginTop + boxHeight + style.MarginBottom;
        return new HtmlRenderFlowBlock(containingWidth, outerHeight, visuals, style.BreakBefore, style.BreakAfter, style.AvoidBreakInside, sourceDescription, pageName: style.PageName);
    }
}

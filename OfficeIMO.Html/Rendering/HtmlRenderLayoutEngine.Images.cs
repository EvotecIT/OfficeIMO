using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock LayoutImage(IElement element, double containingWidth, HtmlRenderBoxStyle style, string? inheritedLink = null) {
        string sourceDescription = HtmlRenderStyleResolver.DescribeSource(element);
        IReadOnlyList<string> candidates = HtmlImageSourceResolver.ResolveImageSourceCandidatesForRendering(element, _baseUri, _resourceUrlPolicy, _options);
        string? source = candidates.FirstOrDefault() ?? element.GetAttribute("src");
        byte[]? bytes = null;
        string contentType = string.Empty;
        OfficeImageInfo? imageInfo = null;
        foreach (string candidate in candidates) {
            if (TryResolveImageSource(candidate, sourceDescription, out bytes, out contentType, out imageInfo, reportDiagnostics: false)) {
                source = candidate;
                break;
            }
        }
        if (bytes == null) TryResolveImageSource(source, sourceDescription, out bytes, out contentType, out imageInfo);
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
        string? link = inheritedLink ?? (element.ParentElement != null && string.Equals(element.ParentElement.TagName, "a", StringComparison.OrdinalIgnoreCase)
            ? ResolveSafeLink(element.ParentElement.GetAttribute("href"), element.ParentElement)
            : null);
        string? alternativeText = element.GetAttribute("alt");
        ReplacedObjectPlacement placement = ResolveReplacedObjectPlacement(
            style,
            contentSize.Width,
            contentSize.Height,
            intrinsicWidth,
            intrinsicHeight);
        bool addedObject = false;
        if (bytes != null && bytes.Length > 0 && placement.IsVisible) {
            if (string.Equals(contentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                if (TryReadSvgDrawing(bytes, sourceDescription, out OfficeDrawing? svgDrawing) && svgDrawing != null) {
                    AddSvgImageVisual(objectVisuals, svgDrawing, imageX, imageY, placement, alternativeText, link, sourceDescription);
                    addedObject = true;
                }
            } else {
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
                addedObject = true;
            }
        }
        if (!addedObject && placement.IsVisible) {
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
        if (!style.PaintVisible) visuals.Clear();

        double outerHeight = style.MarginTop + boxHeight + style.MarginBottom;
        return new HtmlRenderFlowBlock(containingWidth, outerHeight, visuals, style.BreakBefore, style.BreakAfter, style.AvoidBreakInside, sourceDescription, pageName: style.PageName);
    }

    private double ResolveFloatingImageOuterWidth(IElement element, HtmlRenderBoxStyle style) {
        return Math.Max(1D, style.MarginLeft + ResolveReplacedImageBoxWidth(element, style) + style.MarginRight);
    }

    private bool TryReadSvgDrawing(byte[] bytes, string sourceDescription, out OfficeDrawing? drawing) {
        if (OfficeSvgDrawingReader.TryRead(bytes, out drawing, out int unsupportedFeatures) && drawing != null) {
            if (unsupportedFeatures > 0) {
                _diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.SvgContentUnsupported,
                    "Unsupported SVG content was omitted while supported vector content remained active.",
                    HtmlDiagnosticSeverity.Warning,
                    sourceDescription,
                    "features=" + unsupportedFeatures,
                    HtmlConversionLossKind.Omission);
            }
            return true;
        }

        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.SvgContentUnsupported,
            "The SVG image could not be interpreted as a bounded shared vector scene.",
            HtmlDiagnosticSeverity.Warning,
            sourceDescription,
            "image/svg+xml",
            HtmlConversionLossKind.Omission);
        return false;
    }

    private static void AddSvgImageVisual(
        ICollection<HtmlRenderVisual> visuals,
        OfficeDrawing drawing,
        double imageX,
        double imageY,
        ReplacedObjectPlacement placement,
        string? alternativeText,
        string? link,
        string sourceDescription) {
        double visibleX = imageX + placement.X;
        double visibleY = imageY + placement.Y;
        if (!placement.SourceCrop.HasCrop) {
            visuals.Add(new HtmlRenderDrawing(
                drawing,
                visibleX,
                visibleY,
                placement.Width,
                placement.Height,
                visuals.Count,
                alternativeText,
                link,
                sourceDescription));
            return;
        }

        double visibleWidthRatio = Math.Max(
            OfficeImageSourceCrop.MinimumVisibleRatio,
            1D - placement.SourceCrop.Left - placement.SourceCrop.Right);
        double visibleHeightRatio = Math.Max(
            OfficeImageSourceCrop.MinimumVisibleRatio,
            1D - placement.SourceCrop.Top - placement.SourceCrop.Bottom);
        double fullWidth = placement.Width / visibleWidthRatio;
        double fullHeight = placement.Height / visibleHeightRatio;
        var child = new HtmlRenderDrawing(
            drawing,
            visibleX - fullWidth * placement.SourceCrop.Left,
            visibleY - fullHeight * placement.SourceCrop.Top,
            fullWidth,
            fullHeight,
            0,
            alternativeText,
            link,
            sourceDescription);
        visuals.Add(new HtmlRenderClipGroup(
            visibleX,
            visibleY,
            placement.Width,
            placement.Height,
            clipHorizontal: true,
            clipVertical: true,
            new[] { child },
            visuals.Count,
            sourceDescription + ":object-fit-clip"));
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
            bytes = resolvedResource.EncodedBytes;
            contentType = NormalizeImageContentType(resolvedResource.ContentType);
            extension = OfficeImageInfo.GetDefaultExtension(OfficeImageInfo.FromMimeType(contentType));
        } else if (resolvedSource.StartsWith("data:", StringComparison.OrdinalIgnoreCase)
            && HtmlImageDataUri.TryParse(resolvedSource, out HtmlImageDataUri dataUri)) {
            long estimatedBytes;
            try {
                estimatedBytes = dataUri.EstimateDecodedByteCount();
            } catch (FormatException) {
                estimatedBytes = -1L;
            }

            string diagnosticCode = string.Empty;
            string diagnosticDetail = string.Empty;
            bool withinBudget = estimatedBytes >= 0L
                && _resources.CanAcceptInlineResource(estimatedBytes, _options, out diagnosticCode, out diagnosticDetail);
            if (withinBudget && dataUri.TryDecodeBytes(out byte[] decoded)) {
                var inlineResource = new HtmlResolvedResource(decoded, dataUri.MediaType);
                _resources.AddInline(resolvedSource, inlineResource);
                bytes = inlineResource.EncodedBytes;
                contentType = NormalizeImageContentType(inlineResource.ContentType);
                extension = dataUri.FileExtension;
            } else if (!withinBudget && diagnosticCode.Length > 0) {
                _diagnostics.Add(ComponentName, diagnosticCode, "An image data URI exceeded the configured operation-wide resource budget.", HtmlDiagnosticSeverity.Warning, sourceDescription, diagnosticDetail);
            }
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

    private static string NormalizeImageContentType(string contentType) =>
        OfficeImageInfo.TryNormalizeImageContentType(contentType, out string normalized) ? normalized : contentType.Split(';')[0].Trim();
}

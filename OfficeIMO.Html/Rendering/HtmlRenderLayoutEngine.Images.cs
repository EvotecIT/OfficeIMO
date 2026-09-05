using AngleSharp.Dom;
using OfficeIMO.Drawing;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock LayoutImage(IElement element, double containingWidth, HtmlRenderBoxStyle style, string? inheritedLink = null) {
        string? editableImageKey = HtmlEditableLayoutProjector.GetImageSourceKey(element);
        string sourceDescription = string.IsNullOrWhiteSpace(editableImageKey)
            ? HtmlRenderStyleResolver.DescribeSource(element)
            : HtmlEditableLayoutProjector.DescribeImageSource(editableImageKey);
        IReadOnlyList<string> candidates = IsInlineSvgElement(element)
            ? Array.Empty<string>()
            : HtmlImageSourceResolver.ResolveImageSourceCandidatesForRendering(element, _baseUri, _resourceUrlPolicy, _options);
        string? source = candidates.FirstOrDefault() ?? element.GetAttribute("src");
        byte[]? bytes = null;
        string contentType = string.Empty;
        OfficeImageInfo? imageInfo = null;
        if (TryReadInlineSvgSource(element, out byte[]? inlineSvg, out OfficeImageInfo? inlineSvgInfo)) {
            bytes = inlineSvg;
            contentType = "image/svg+xml";
            imageInfo = inlineSvgInfo;
        } else {
            foreach (string candidate in candidates) {
                if (TryResolveImageSource(candidate, sourceDescription, out bytes, out contentType, out imageInfo, reportDiagnostics: false)) {
                    source = candidate;
                    break;
                }
            }
            if (bytes == null) TryResolveImageSource(source, sourceDescription, out bytes, out contentType, out imageInfo);
        }
        if (bytes != null
            && OfficeImageOrientationNormalizer.TryNormalizeToPng(
                bytes,
                style.ApplyEmbeddedImageOrientation,
                out byte[] orientedPng,
                out OfficeImageInfo? orientedInfo)) {
            bytes = orientedPng;
            contentType = "image/png";
            imageInfo = orientedInfo;
        }
        bool hasIntrinsicSize = imageInfo != null && imageInfo.Width > 0 && imageInfo.Height > 0;
        double intrinsicWidth = hasIntrinsicSize
            ? imageInfo!.Width * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, style.ImageResolutionDpi ?? imageInfo.DpiX)
            : 300D;
        double intrinsicHeight = hasIntrinsicSize
            ? imageInfo!.Height * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, style.ImageResolutionDpi ?? imageInfo.DpiY)
            : 150D;
        ReplacedContentSize contentSize = ResolveReplacedContentSize(style, intrinsicWidth, intrinsicHeight, hasIntrinsicSize);
        double boxWidth = contentSize.Width + style.HorizontalInsets;
        double boxHeight = contentSize.Height + style.VerticalInsets;
        EnsureReplacedBoxSize(boxWidth, boxHeight);
        var visuals = new List<HtmlRenderVisual>();
        var objectVisuals = new List<HtmlRenderVisual>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        double imageX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double imageY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        string? link = inheritedLink ?? (element.ParentElement != null && string.Equals(element.ParentElement.TagName, "a", StringComparison.OrdinalIgnoreCase)
            ? ResolveSafeLink(element.ParentElement.GetAttribute("href"), element.ParentElement)
            : null);
        string? alternativeText = element.GetAttribute("alt") ?? element.GetAttribute("aria-label");
        ReplacedObjectPlacement placement = ResolveReplacedObjectPlacement(
            style,
            contentSize.Width,
            contentSize.Height,
            intrinsicWidth,
            intrinsicHeight);
        bool addedObject = false;
        if (bytes != null && bytes.Length > 0 && placement.IsVisible) {
            if (string.Equals(contentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                if (TryReadSvgDrawing(
                    bytes,
                    hasIntrinsicSize ? intrinsicWidth : 0D,
                    hasIntrinsicSize ? intrinsicHeight : 0D,
                    sourceDescription,
                    out OfficeDrawing? svgDrawing) && svgDrawing != null) {
                    AddSvgImageVisual(objectVisuals, svgDrawing, bytes, imageX, imageY, placement, alternativeText,
                        link, sourceDescription);
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
                objectVisuals.Add(new HtmlRenderText(alternativeText!, imageX + placement.X + 4D, imageY + placement.Y + 4D, Math.Max(1D, placement.Width - 8D), Math.Max(1D, textHeight), style.Font, style.Color, OfficeTextAlignment.Left, style.LineHeight, objectVisuals.Count, link, sourceDescription, "figure-alternative-text", null, null, null, false, null, null, style.UnderlineStyle, style.StrikethroughStyle, style.Baseline, style.BaselineLevel, style.BaselineScale, style.BaselineOffset, decorationColor: style.DecorationColor, featureSettings: style.TextFeatureSettings, fontPalette: style.FontPalette));
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

    private static bool IsReplacedImageElement(IElement element) =>
        IsReplacedImageElementTag(element.LocalName);

    private static bool IsReplacedImageElementTag(string tagName) =>
        tagName.Equals("img", StringComparison.OrdinalIgnoreCase)
        || tagName.Equals("svg", StringComparison.OrdinalIgnoreCase);

    private static bool IsInlineSvgElement(IElement element) =>
        element.LocalName.Equals("svg", StringComparison.OrdinalIgnoreCase);

    private bool TryReadInlineSvgSource(
        IElement element,
        out byte[]? bytes,
        out OfficeImageInfo? imageInfo) {
        bytes = null;
        imageInfo = null;
        if (!IsInlineSvgElement(element)) return false;

        string source = element.OuterHtml;
        try {
            XElement svg = XElement.Parse(source, LoadOptions.PreserveWhitespace);
            IReadOnlyList<IElement> htmlElements = new[] { element }
                .Concat(element.QuerySelectorAll("*").OfType<IElement>())
                .ToArray();
            IReadOnlyList<XElement> svgElements = svg.DescendantsAndSelf().ToArray();
            int count = Math.Min(htmlElements.Count, svgElements.Count);
            for (int index = 0; index < count; index++) {
                if (!_computedStyles.Elements.TryGetValue(htmlElements[index], out HtmlComputedStyle? computed)) continue;
                string computedSvgStyle = BuildInlineSvgComputedStyle(computed);
                if (computedSvgStyle.Length > 0) svgElements[index].SetAttributeValue("style", computedSvgStyle);
            }
            source = svg.ToString(SaveOptions.DisableFormatting);
        } catch (System.Xml.XmlException) {
            // The bounded SVG reader reports malformed XML through its normal parse result.
        }
        bytes = Encoding.UTF8.GetBytes(source);
        if (OfficeImageReader.TryIdentify(bytes, ".svg", out OfficeImageInfo identified)) {
            imageInfo = identified;
        }
        return true;
    }

    private static string BuildInlineSvgComputedStyle(HtmlComputedStyle computed) {
        var style = new StringBuilder();
        foreach (KeyValuePair<string, string> property in computed.Properties) {
            if (!property.Key.StartsWith("--", StringComparison.Ordinal)
                && !IsSvgComputedStyleProperty(property.Key)) continue;
            if (style.Length > 0) style.Append(';');
            style.Append(property.Key).Append(':').Append(property.Value);
        }
        return style.ToString();
    }

    private static bool IsSvgComputedStyleProperty(string name) => name.ToLowerInvariant() is
        "color" or "fill" or "fill-opacity" or "fill-rule" or "stroke" or "stroke-width"
        or "stroke-opacity" or "stroke-dasharray" or "stroke-dashoffset" or "stroke-linecap"
        or "stroke-linejoin" or "stroke-miterlimit" or "opacity" or "display" or "visibility"
        or "font-family" or "font-size" or "font-style" or "font-weight" or "line-height"
        or "writing-mode" or "text-orientation" or "text-anchor" or "dominant-baseline"
        or "baseline-shift" or "transform" or "clip-path" or "filter" or "mask"
        or "mix-blend-mode" or "marker-start" or "marker-mid" or "marker-end";

    private bool TryReadSvgDrawing(
        byte[] bytes,
        double fallbackWidth,
        double fallbackHeight,
        string sourceDescription,
        out OfficeDrawing? drawing) {
        var readerOptions = new OfficeSvgDrawingReaderOptions();
        readerOptions.Fonts.AddRange(_fonts);
        if (_options.SvgForeignObjectDepth < _options.MaxSvgForeignObjectDepth) {
            readerOptions.ForeignObjectRenderer = RenderSvgForeignObject;
        }
        if (OfficeSvgDrawingReader.TryRead(bytes, readerOptions, out drawing, out int unsupportedFeatures) && drawing != null) {
            if (unsupportedFeatures > 0) {
                if (TryRasterizeSvgFallback(bytes, drawing.Width, drawing.Height, sourceDescription, unsupportedFeatures, out OfficeDrawing? rasterFallback)) {
                    drawing = rasterFallback;
                    return true;
                }
                _diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.SvgContentUnsupported,
                    "Unsupported SVG content was omitted while supported vector content remained active.",
                    HtmlDiagnosticSeverity.Warning,
                    sourceDescription,
                    "features=" + unsupportedFeatures,
                    OfficeConversionLossKind.Omission);
            }
            return true;
        }

        if (TryRasterizeSvgFallback(bytes, fallbackWidth, fallbackHeight, sourceDescription, null, out drawing)) {
            return true;
        }

        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.SvgContentUnsupported,
            "The SVG image could not be interpreted as a bounded shared vector scene.",
            HtmlDiagnosticSeverity.Warning,
            sourceDescription,
            "image/svg+xml",
            OfficeConversionLossKind.Omission);
        return false;
    }

    private OfficeDrawing? RenderSvgForeignObject(OfficeSvgForeignObjectContext context) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (_options.SvgForeignObjectDepth >= _options.MaxSvgForeignObjectDepth) return null;

        string width = context.Width.ToString("0.################", System.Globalization.CultureInfo.InvariantCulture);
        string height = context.Height.ToString("0.################", System.Globalization.CultureInfo.InvariantCulture);
        string source = "<!doctype html><html><head><meta charset='utf-8'><style>"
            + "html,body{margin:0;padding:0;width:" + width + "px;height:" + height + "px;overflow:hidden;background:transparent}"
            + "</style></head><body>" + context.Html + "</body></html>";
        HtmlConversionLimits limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxHtmlNodes = Math.Min(limits.MaxHtmlNodes ?? int.MaxValue, _options.MaxSvgForeignObjectHtmlNodes);
        limits.MaxHtmlDepth = Math.Min(limits.MaxHtmlDepth ?? int.MaxValue, _options.MaxLayoutDepth);
        var conversionOptions = new HtmlConversionDocumentOptions {
            BaseUri = _baseUri,
            UrlPolicy = _options.UrlPolicy.Clone(),
            ResourceUrlPolicy = _options.GetResourceUrlPolicy().Clone(),
            Limits = limits,
            IncludeNormalizedHtml = false,
            UseBodyContentsOnly = true
        };
        HtmlConversionDocument nestedDocument = HtmlConversionDocument.Parse(source, conversionOptions);
        HtmlRenderOptions nestedOptions = _options.Clone();
        nestedOptions.Mode = HtmlRenderMode.Continuous;
        nestedOptions.ViewportWidth = context.Width;
        nestedOptions.ViewportHeight = context.Height;
        nestedOptions.Margins = HtmlRenderMargins.All(0D);
        nestedOptions.HonorCssPageRules = false;
        nestedOptions.BackgroundColor = OfficeColor.Transparent;
        nestedOptions.FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss;
        nestedOptions.MaxHtmlNodes = Math.Min(nestedOptions.MaxHtmlNodes, _options.MaxSvgForeignObjectHtmlNodes);
        nestedOptions.SvgForeignObjectDepth = _options.SvgForeignObjectDepth + 1;
        nestedOptions.ResourceResolver = null;
        nestedOptions.SynchronousResourceResolver = null;
        nestedOptions.AdditionalStylesheets.Clear();

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(nestedDocument, nestedOptions, _cancellationToken);
        _diagnostics.AddRange(rendered.Diagnostics);
        HtmlRenderPage page = rendered.Pages[0];
        OfficeDrawing nested = page.CreateDrawing(_cancellationToken);
        if (Math.Abs(nested.Width - context.Width) <= 0.0001D
            && Math.Abs(nested.Height - context.Height) <= 0.0001D) return nested;

        var clipped = new OfficeDrawing(context.Width, context.Height);
        clipped.AddClippedDrawing(
            nested,
            0D,
            0D,
            OfficeClipPath.Rectangle(context.Width, context.Height));
        return clipped;
    }

    private bool TryRasterizeSvgFallback(
        byte[] bytes,
        double width,
        double height,
        string sourceDescription,
        int? unsupportedFeatures,
        out OfficeDrawing? drawing) {
        drawing = null;
        if (_options.ImageCodec == null) return false;
        try {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!_options.ImageCodec.TryDecode(bytes, "image/svg+xml", out OfficeRasterImage? raster) || raster == null) return false;
            _cancellationToken.ThrowIfCancellationRequested();
            long pixels = checked((long)raster.Width * raster.Height);
            if (pixels > _options.MaximumRasterPixels
                || raster.Width > _options.MaxSurfaceWidth
                || raster.Height > _options.MaxSurfaceHeight) return false;
            byte[] png = OfficeRasterImageEncoder.Encode(raster, OfficeImageExportFormat.Png, _options.RasterEncoding);
            _cancellationToken.ThrowIfCancellationRequested();
            double resolvedWidth = width > 0D ? width : raster.Width;
            double resolvedHeight = height > 0D ? height : raster.Height;
            var fallback = new OfficeDrawing(resolvedWidth, resolvedHeight);
            fallback.AddImage(
                png,
                "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, resolvedWidth, resolvedHeight)));
            drawing = fallback;
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.SvgRasterFallback,
                "A caller-supplied codec rasterized unsupported SVG features.",
                HtmlDiagnosticSeverity.Info,
                sourceDescription,
                "features=" + (unsupportedFeatures.HasValue
                    ? unsupportedFeatures.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : "managed-parse-failed") + ";pixels=" + pixels,
                OfficeConversionLossKind.Approximation);
            return true;
        } catch (Exception exception) when (exception is not OperationCanceledException && exception is not OutOfMemoryException && exception is not StackOverflowException) {
            return false;
        }
    }

    private static void AddSvgImageVisual(
        ICollection<HtmlRenderVisual> visuals,
        OfficeDrawing drawing,
        byte[] imageBytes,
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
                sourceDescription,
                imageBytes: imageBytes,
                imageContentType: "image/svg+xml",
                sourceCrop: placement.SourceCrop));
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
            sourceDescription,
            imageBytes: imageBytes,
            imageContentType: "image/svg+xml",
            sourceCrop: placement.SourceCrop,
            imageX: visibleX,
            imageY: visibleY,
            imageWidth: placement.Width,
            imageHeight: placement.Height);
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
                && _resources.CanAcceptInlineResource(estimatedBytes, out diagnosticCode, out diagnosticDetail);
            if (withinBudget && dataUri.TryDecodeBytes(out byte[] decoded)) {
                var inlineResource = new HtmlResolvedResource(decoded, dataUri.MediaType);
                if (_resources.TryAcceptInline(HtmlResourceKind.Image, resolvedSource, inlineResource,
                        out diagnosticCode, out diagnosticDetail)) {
                    bytes = inlineResource.EncodedBytes;
                    contentType = NormalizeImageContentType(inlineResource.ContentType);
                    extension = dataUri.FileExtension;
                } else if (diagnosticCode.Length > 0) {
                    _diagnostics.Add(ComponentName, diagnosticCode,
                        "An image data URI was rejected by the shared resource session.",
                        HtmlDiagnosticSeverity.Warning, sourceDescription, diagnosticDetail);
                }
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

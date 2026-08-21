namespace OfficeIMO.Html;

public static partial class HtmlRtfConverterExtensions {
    /// <summary>Imports a prepared shared HTML document into RTF and returns structured evidence.</summary>
    public static HtmlToRtfResult ToRtfDocumentResult(this HtmlConversionDocument document, HtmlToRtfOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlToRtfOptions resolved = (options ?? new HtmlToRtfOptions()).Clone();
        HtmlUrlPolicy requestedHyperlinkPolicy = resolved.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile();
        HtmlUrlPolicy requestedResourcePolicy = resolved.ResourceUrlPolicy ?? requestedHyperlinkPolicy;
        resolved.UrlPolicy = HtmlUrlPolicy.Intersect(document.HyperlinkUrlPolicy, requestedHyperlinkPolicy);
        resolved.ResourceUrlPolicy = HtmlUrlPolicy.Intersect(document.ResourceUrlPolicy, requestedResourcePolicy);
        if (resolved.BaseUri == null) resolved.BaseUri = document.BaseUri;
        HtmlEditableLayoutRegionKinds regionKinds =
            HtmlEditableLayoutRegionKinds.Positioned | HtmlEditableLayoutRegionKinds.Floating;
        HtmlEditableLayoutProjection? editableLayout = resolved.ImportEditableLayoutRegions
            && HtmlEditableLayoutProjector.MayContainEditableLayoutRegions(document, regionKinds)
            ? HtmlEditableLayoutProjector.Project(
                document,
                mediaContext: document.MediaContext,
                regionKinds: regionKinds,
                maximumEditableSurfaceNumber: document.MediaContext == HtmlCssMediaContext.Print ? 0 : 1)
            : null;
        AddEditableLayoutDiagnostics(editableLayout, resolved);
        AngleSharp.Html.Dom.IHtmlDocument sourceDocument = editableLayout?.RemainingDocument
            ?? document.CreateSourceDocumentForConversion();
        HtmlNormalizer.SanitizePreparedDocumentStructure(sourceDocument);
        HtmlActiveMediaFilter.Filter(sourceDocument, document.MediaContext);
        RtfDocument rtfDocument = RtfHtmlReader.Read(sourceDocument, resolved);
        if (editableLayout?.Regions.Count > 0) AddEditableLayoutFrames(rtfDocument, editableLayout, resolved);
        return new HtmlToRtfResult(
            rtfDocument,
            document.Diagnostics.Concat(document.ResourceManifest.Diagnostics).Concat(resolved.HtmlDiagnostics),
            resolved.Diagnostics.AsReadOnly(),
            resolved.ConversionReport);
    }

    private static void AddEditableLayoutDiagnostics(HtmlEditableLayoutProjection? projection, HtmlToRtfOptions options) {
        if (projection == null) return;
        foreach (HtmlDiagnostic diagnostic in projection.Diagnostics) {
            options.AddDiagnostic(new HtmlRtfConversionDiagnostic(
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.Severity == HtmlDiagnosticSeverity.Error
                    ? HtmlRtfConversionDiagnosticSeverity.Error
                    : diagnostic.Severity == HtmlDiagnosticSeverity.Warning
                        ? HtmlRtfConversionDiagnosticSeverity.Warning
                        : HtmlRtfConversionDiagnosticSeverity.Info,
                diagnostic.Source,
                diagnostic.Detail,
                diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected
                    ? RtfConversionAction.Preserved
                    : RtfConversionAction.Substituted));
        }
    }

    private static void AddEditableLayoutFrames(
        RtfDocument document,
        HtmlEditableLayoutProjection projection,
        HtmlToRtfOptions options) {
        const double twipsPerCssPixel = 15D;
        foreach (HtmlRenderLayoutRegion region in projection.Regions.OrderBy(item => item.PaintOrder)) {
            RtfParagraph paragraph = document.AddParagraph(region.SourceText);
            int horizontalPosition = ToBoundedFrameCoordinate(
                region.X, twipsPerCssPixel, out bool horizontalSimplified);
            int verticalPosition = ToBoundedFrameCoordinate(
                region.Y, twipsPerCssPixel, out bool verticalSimplified);
            int nativeWidth = ToBoundedFrameSize(
                region.Width, twipsPerCssPixel, out bool widthSimplified);
            int nativeHeight = ToBoundedFrameSize(
                region.Height, twipsPerCssPixel, out bool heightSimplified);
            paragraph.Frame
                .SetSize(
                    nativeWidth,
                    -nativeHeight)
                .SetAnchors(RtfParagraphFrameHorizontalAnchor.Page, RtfParagraphFrameVerticalAnchor.Page)
                .SetPosition(
                    RtfParagraphFrameHorizontalPosition.Absolute,
                    horizontalPosition,
                    RtfParagraphFrameVerticalPosition.Absolute,
                    verticalPosition)
                .SetWrapping(
                    noWrap: region.RegionKind == HtmlRenderLayoutRegionKind.Positioned,
                    overlayText: region.RegionKind == HtmlRenderLayoutRegionKind.Positioned,
                    noOverlap: region.RegionKind == HtmlRenderLayoutRegionKind.Floating);
            if (horizontalSimplified || verticalSimplified || widthSimplified || heightSimplified) {
                options.AddDiagnostic(HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                    "RTF bounded an editable layout frame's page position or size to its native range.",
                    region.Source, severity: HtmlRtfConversionDiagnosticSeverity.Warning,
                    action: RtfConversionAction.Substituted);
            }
            if (region.BackgroundColor.HasValue) {
                paragraph.SetBackgroundColor(document.AddColor(
                    region.BackgroundColor.Value.R,
                    region.BackgroundColor.Value.G,
                    region.BackgroundColor.Value.B));
            }
            IReadOnlyList<AngleSharp.Html.Dom.IHtmlImageElement> sourceImages = projection.GetSourceImages(region);
            IReadOnlyList<(HtmlRenderImage Image, double Opacity)> renderedImages =
                HtmlEditableLayoutProjector.EnumerateImages(region.Visuals, includeBackgroundImages: false)
                    .ToList()
                    .AsReadOnly();
            var projectedSources = new HashSet<AngleSharp.Html.Dom.IHtmlImageElement>();
            foreach ((HtmlRenderImage visual, double opacity) in renderedImages) {
                AngleSharp.Html.Dom.IHtmlImageElement? sourceImage = projection.GetSourceImage(visual);
                if (sourceImage == null || !projectedSources.Add(sourceImage)) continue;
                AddRtfRegionImage(paragraph, sourceImage, visual, opacity, region, options, twipsPerCssPixel);
            }
            foreach (AngleSharp.Html.Dom.IHtmlImageElement sourceImage in sourceImages) {
                if (projectedSources.Add(sourceImage)) {
                    AddRtfRegionImage(paragraph, sourceImage, null, 1D, region, options, twipsPerCssPixel);
                }
            }
            if (region.BackgroundLayerCount > 0) {
                options.AddDiagnostic(HtmlEditableLayoutDiagnosticCodes.BackgroundLayersFlattened,
                    "RTF retained the solid editable frame background; extra CSS background image layers were omitted.",
                    region.Source, severity: HtmlRtfConversionDiagnosticSeverity.Warning,
                    action: RtfConversionAction.Flattened);
            }
            if (region.BoxShadowLayerCount > 0 || region.ZIndex != 0) {
                options.AddDiagnostic(HtmlEditableLayoutDiagnosticCodes.EffectUnsupported,
                    "RTF retained editable frame geometry and wrap behavior without unsupported CSS shadow or explicit stacking metadata.",
                    region.Source, severity: HtmlRtfConversionDiagnosticSeverity.Warning,
                    action: RtfConversionAction.Substituted);
            }
        }
    }

    private static void AddRtfRegionImage(
        RtfParagraph paragraph,
        AngleSharp.Html.Dom.IHtmlImageElement sourceImage,
        HtmlRenderImage? visual,
        double opacity,
        HtmlRenderLayoutRegion region,
        HtmlToRtfOptions options,
        double twipsPerCssPixel) {
        string source = HtmlImageSourceResolver.ResolveImageSource(
            sourceImage, options.BaseUri, options.GetResourceUrlPolicy());
        if (!TryReadRtfRegionImage(source, out RtfImageFormat format, out byte[] bytes, out string detail)) {
            options.AddDiagnostic(new HtmlRtfConversionDiagnostic(
                HtmlEditableLayoutDiagnosticCodes.RegionImageOmitted,
                "A picture inside an editable RTF layout region was omitted.",
                HtmlRtfConversionDiagnosticSeverity.Warning,
                string.IsNullOrWhiteSpace(source) ? region.Source : source,
                detail,
                RtfConversionAction.Omitted));
            return;
        }
        RtfImage nativeImage = paragraph.AddImage(format, bytes);
        nativeImage.Description = sourceImage.AlternativeText;
        if (visual == null) return;
        nativeImage.DesiredWidthTwips = checked((int)Math.Round(visual.Width * twipsPerCssPixel));
        nativeImage.DesiredHeightTwips = checked((int)Math.Round(visual.Height * twipsPerCssPixel));
        if (opacity < 0.999D || visual.SourceCrop.HasCrop) {
            options.AddDiagnostic(new HtmlRtfConversionDiagnostic(
                HtmlEditableLayoutDiagnosticCodes.EffectUnsupported,
                "RTF retained an editable region picture without unsupported CSS alpha or source-crop effects.",
                HtmlRtfConversionDiagnosticSeverity.Warning,
                visual.Source,
                "opacity=" + opacity.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)
                    + "; crop=" + visual.SourceCrop.HasCrop,
                RtfConversionAction.Substituted));
        }
    }

    private static int ToBoundedFrameCoordinate(double cssPixels, double unitsPerCssPixel, out bool simplified) {
        double value = Math.Round(cssPixels * unitsPerCssPixel);
        if (double.IsNaN(value)) {
            simplified = true;
            return 0;
        }
        if (value <= int.MinValue) {
            simplified = value < int.MinValue;
            return int.MinValue;
        }
        if (value >= int.MaxValue) {
            simplified = value > int.MaxValue;
            return int.MaxValue;
        }
        simplified = false;
        return (int)value;
    }

    private static int ToBoundedFrameSize(double cssPixels, double unitsPerCssPixel, out bool simplified) {
        double value = Math.Round(cssPixels * unitsPerCssPixel);
        if (double.IsNaN(value) || value <= 1D) {
            simplified = value != 1D;
            return 1;
        }
        if (value >= int.MaxValue) {
            simplified = value > int.MaxValue;
            return int.MaxValue;
        }
        simplified = false;
        return (int)value;
    }

    private static bool TryReadRtfRegionImage(
        string source,
        out RtfImageFormat format,
        out byte[] bytes,
        out string detail) {
        format = RtfImageFormat.Unknown;
        bytes = Array.Empty<byte>();
        detail = string.Empty;
        if (!HtmlImageDataUri.TryParse(source, out HtmlImageDataUri dataUri) || !dataUri.IsBase64) {
            detail = "RTF layout-region pictures must use an embedded base64 PNG or JPEG data URI.";
            return false;
        }
        if (dataUri.MediaType.Equals("image/png", StringComparison.OrdinalIgnoreCase)) {
            format = RtfImageFormat.Png;
        } else if (dataUri.MediaType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase)
                   || dataUri.MediaType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase)) {
            format = RtfImageFormat.Jpeg;
        } else {
            detail = "The RTF editable frame supports PNG and JPEG picture payloads.";
            return false;
        }
        if (!dataUri.TryDecodeBytes(out bytes) || bytes.Length == 0) {
            detail = "The embedded layout-region picture payload could not be decoded.";
            format = RtfImageFormat.Unknown;
            bytes = Array.Empty<byte>();
            return false;
        }
        return true;
    }

    /// <summary>Exports RTF to semantic HTML and returns structured conversion evidence.</summary>
    public static RtfToHtmlResult ToHtmlResult(this RtfDocument document, RtfToHtmlOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        RtfToHtmlOptions resolved = (options ?? new RtfToHtmlOptions()).Clone();
        string html = ToHtmlCore(document, resolved);
        return new RtfToHtmlResult(html, resolved.HtmlDiagnostics, resolved.Diagnostics.AsReadOnly(), resolved.ConversionReport);
    }
}

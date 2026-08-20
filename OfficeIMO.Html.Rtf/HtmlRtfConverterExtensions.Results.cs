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
                regionKinds: regionKinds)
            : null;
        AddEditableLayoutDiagnostics(editableLayout, resolved);
        AngleSharp.Html.Dom.IHtmlDocument sourceDocument = editableLayout?.RemainingDocument
            ?? document.CreateSourceDocumentForConversion();
        HtmlNormalizer.SanitizePreparedDocumentStructure(sourceDocument);
        HtmlActiveMediaFilter.Filter(sourceDocument, document.MediaContext);
        RtfDocument rtfDocument = RtfHtmlReader.Read(sourceDocument, resolved);
        if (editableLayout?.Regions.Count > 0) AddEditableLayoutFrames(rtfDocument, editableLayout.Regions, resolved);
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
        IReadOnlyList<HtmlRenderLayoutRegion> regions,
        HtmlToRtfOptions options) {
        const double twipsPerCssPixel = 15D;
        foreach (HtmlRenderLayoutRegion region in regions.OrderBy(item => item.PaintOrder)) {
            RtfParagraph paragraph = document.AddParagraph(region.SourceText);
            paragraph.Frame
                .SetSize(
                    checked((int)Math.Round(region.Width * twipsPerCssPixel)),
                    -checked((int)Math.Round(region.Height * twipsPerCssPixel)))
                .SetAnchors(RtfParagraphFrameHorizontalAnchor.Page, RtfParagraphFrameVerticalAnchor.Page)
                .SetPosition(
                    RtfParagraphFrameHorizontalPosition.Absolute,
                    checked((int)Math.Round(region.X * twipsPerCssPixel)),
                    RtfParagraphFrameVerticalPosition.Absolute,
                    checked((int)Math.Round(region.Y * twipsPerCssPixel)))
                .SetWrapping(
                    noWrap: region.RegionKind == HtmlRenderLayoutRegionKind.Positioned,
                    overlayText: region.RegionKind == HtmlRenderLayoutRegionKind.Positioned,
                    noOverlap: region.RegionKind == HtmlRenderLayoutRegionKind.Floating);
            if (region.BackgroundColor.HasValue) {
                paragraph.SetBackgroundColor(document.AddColor(
                    region.BackgroundColor.Value.R,
                    region.BackgroundColor.Value.G,
                    region.BackgroundColor.Value.B));
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

    /// <summary>Exports RTF to semantic HTML and returns structured conversion evidence.</summary>
    public static RtfToHtmlResult ToHtmlResult(this RtfDocument document, RtfToHtmlOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        RtfToHtmlOptions resolved = (options ?? new RtfToHtmlOptions()).Clone();
        string html = ToHtmlCore(document, resolved);
        return new RtfToHtmlResult(html, resolved.HtmlDiagnostics, resolved.Diagnostics.AsReadOnly(), resolved.ConversionReport);
    }
}

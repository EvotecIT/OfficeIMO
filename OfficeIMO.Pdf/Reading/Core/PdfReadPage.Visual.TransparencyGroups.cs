using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private bool IsSupportedType3TransparencyGroup(PdfDictionary formDictionary) {
        if (ResolveEffectObject(formDictionary.Items.TryGetValue("Group", out PdfObject? groupObject) ? groupObject : null) is not PdfDictionary group ||
            ResolveEffectObject(group.Items.TryGetValue("S", out PdfObject? subtypeObject) ? subtypeObject : null) is not PdfName { Name: "Transparency" } ||
            ResolveEffectObject(group.Items.TryGetValue("I", out PdfObject? isolatedObject) ? isolatedObject : null) is not PdfBoolean { Value: true }) {
            return false;
        }
        if (group.Items.TryGetValue("K", out PdfObject? knockoutObject) &&
            ResolveEffectObject(knockoutObject) is not PdfBoolean { Value: false }) {
            return false;
        }
        if (!group.Items.TryGetValue("CS", out PdfObject? colorSpaceObject) ||
            ResolveEffectObject(colorSpaceObject) is not PdfName { Name: "DeviceRGB" }) {
            return false;
        }
        return true;
    }

    private bool TryCreateType3TransparencyGroupDrawing(
        string content,
        PdfDictionary? resources,
        Matrix2D transform,
        double pageWidth,
        double pageHeight,
        PdfPageXObjectInvocation invocation,
        HashSet<PdfStream> activeForms,
        HashSet<PdfStream> activeType3Glyphs,
        HashSet<double> renderedType3PaintOrders,
        Type3GlyphBudget type3GlyphBudget,
        double paintOrderScale,
        bool includeTilingPatterns,
        bool retainPrimitiveData,
        Dictionary<(PdfStream Stream, PdfDictionary Resources), PdfPageTilingPatternResource?>? tilingPatternResourceCache,
        TextContentParser.TextOutputBudget? textOutputBudget,
        PageContentBudget pageContentBudget,
        int contentNestingDepth,
        PdfContentOrderKey? contentOrderPrefix,
        out OfficeDrawing groupDrawing) {
        int failureVersion = type3GlyphBudget.FailureVersion;
        var elements = new List<PdfPageDrawingElement>();
        CollectVisualPrimitivesAndForms(
            content,
            resources,
            transform,
            pageWidth,
            pageHeight,
            primitive => elements.Add(PdfPageDrawingElement.FromPrimitive(primitive, elements.Count)),
            activeForms,
            activeType3Glyphs,
            renderedType3PaintOrders,
            type3GlyphBudget,
            invocation.PaintOrder,
            paintOrderScale * 0.000000001D,
            initialClipPath: invocation.ClipPath,
            initialFillColor: invocation.FillColor,
            initialFillColorSpace: invocation.FillColorSpace,
            initialFillOpacity: invocation.FillOpacity,
            initialStrokeColor: invocation.StrokeColor,
            initialStrokeColorSpace: invocation.StrokeColorSpace,
            initialStrokeOpacity: invocation.StrokeOpacity,
            initialStrokeWidth: invocation.StrokeWidth,
            initialStrokeDashStyle: invocation.StrokeDashStyle,
            initialStrokeLineCap: invocation.StrokeLineCap,
            initialStrokeLineJoin: invocation.StrokeLineJoin,
            contentNestingDepth: contentNestingDepth + 1,
            includeTilingPatterns: includeTilingPatterns,
            retainPrimitiveData: retainPrimitiveData,
            requireSupportedType3Content: true,
            allowSupportedType3Patterns: true,
            allowSupportedType3TransparencyGroups: true,
            type3ImageVisitor: (placement, image, effect) => elements.Add(
                PdfPageDrawingElement.FromImage(placement, image, elements.Count).WithEffect(effect)),
            type3PrimitiveVisitor: (primitive, effect) => elements.Add(
                PdfPageDrawingElement.FromPrimitive(primitive, elements.Count).WithEffect(effect)),
            type3GroupVisitor: (drawing, transform, paintOrder, key, effect) => elements.Add(
                PdfPageDrawingElement.FromGroup(drawing, transform, paintOrder, key, elements.Count).WithEffect(effect)),
            tilingPatternResourceCache: tilingPatternResourceCache,
            textOutputBudget: textOutputBudget,
            pageContentBudget: pageContentBudget,
            contentOrderPrefix: contentOrderPrefix);
        if (type3GlyphBudget.FailureVersion != failureVersion) {
            groupDrawing = null!;
            return false;
        }

        var placements = new List<PdfImagePlacement>();
        CollectImagePlacementsAndForms(
            content,
            resources,
            0,
            transform,
            pageHeight,
            placements,
            activeForms,
            invocation.FillColor,
            invocation.FillColorSpace,
            invocation.FillOpacity,
            invocation.PaintOrder,
            paintOrderScale * 0.000000001D,
            initialClipPath: invocation.ClipPath,
            contentNestingDepth: contentNestingDepth + 1,
            pageContentBudget: pageContentBudget,
            contentOrderPrefix: contentOrderPrefix,
            skipTransparencyGroupForms: true);
        if (placements.Count > 0) {
            IReadOnlyList<PdfExtractedImage> images = GetImagesForResources(resources, 0, placements, colorizeImageMasks: true);
            for (int i = 0; i < placements.Count; i++) {
                PdfExtractedImage? image = FindImage(images, placements[i]);
                if (image == null || !image.IsImageFile) {
                    groupDrawing = null!;
                    return false;
                }
                elements.Add(PdfPageDrawingElement.FromImage(placements[i], image, elements.Count));
            }
        }

        var effects = new List<PdfPageDrawingEffectTransition>();
        CollectGraphicsEffectTransitions(
            content,
            resources,
            transform,
            pageHeight,
            effects,
            new HashSet<PdfStream>(),
            PdfPageDrawingEffect.Default,
            invocation.PaintOrder,
            paintOrderScale * 0.000000001D,
            initialClipPath: invocation.ClipPath,
            initialFillColor: invocation.FillColor,
            initialFillColorSpace: invocation.FillColorSpace,
            initialFillOpacity: invocation.FillOpacity,
            initialStrokeColor: invocation.StrokeColor,
            initialStrokeColorSpace: invocation.StrokeColorSpace,
            initialStrokeOpacity: invocation.StrokeOpacity,
            initialStrokeWidth: invocation.StrokeWidth,
            initialStrokeDashStyle: invocation.StrokeDashStyle,
            initialStrokeLineCap: invocation.StrokeLineCap,
            initialStrokeLineJoin: invocation.StrokeLineJoin,
            contentNestingDepth: contentNestingDepth + 1,
            pageContentBudget: pageContentBudget,
            contentOrderPrefix: contentOrderPrefix);
        SortGraphicsEffectTransitions(effects);
        OverlayDrawingEffects(elements, effects);
        SortDrawingElements(elements);

        groupDrawing = new OfficeDrawing(pageWidth, pageHeight);
        var softMasks = new Dictionary<(PdfStream Group, OfficeSoftMaskMode Mode, OfficeColor Backdrop, Matrix2D Transform, double Width, double Height), OfficeDrawingSoftMask>();
        var activeSoftMasks = new HashSet<PdfStream>();
        TextContentParser.TextOutputBudget outputBudget = textOutputBudget ?? CreateTextOutputBudget();
        for (int i = 0; i < elements.Count; i++) {
            AddDrawingElement(groupDrawing, pageHeight, transform, elements[i], softMasks, activeSoftMasks, outputBudget, pageContentBudget, type3GlyphBudget);
        }
        return true;
    }
}

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

    private Type3TransparencyGroupDrawingResult TryCreateType3TransparencyGroupDrawing(
        string content,
        PdfDictionary formDictionary,
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
        bool requireNestedType3Uncolored,
        Dictionary<(PdfStream Stream, PdfDictionary Resources), PdfPageTilingPatternResource?>? tilingPatternResourceCache,
        TextContentParser.TextOutputBudget? textOutputBudget,
        PageContentBudget pageContentBudget,
        int contentNestingDepth,
        PdfContentOrderKey? contentOrderPrefix,
        out OfficeDrawing groupDrawing,
        out OfficeTransform groupTransform) {
        groupDrawing = null!;
        groupTransform = OfficeTransform.Identity;
        Type3TransparencyGroupDrawingResult boundsResult = TryGetVisibleType3TransparencyGroupBounds(
                formDictionary,
                transform,
                invocation.ClipPath,
                pageWidth,
                pageHeight,
                out PdfPageClipPath fittedBounds);
        if (boundsResult != Type3TransparencyGroupDrawingResult.Success) return boundsResult;

        double localPageWidth = fittedBounds.Width;
        double localPageHeight = fittedBounds.Height;
        Matrix2D localTransform = Matrix2D.Multiply(
            Matrix2D.Translation(
                -fittedBounds.X,
                localPageHeight - pageHeight + fittedBounds.Y),
            transform);
        PdfPageClipPath? localClipPath = fittedBounds.IsRectangle
            ? null
            : fittedBounds.Translate(fittedBounds.X, fittedBounds.Y);
        PdfPagePatternSelection? localFillPattern = invocation.FillPattern?.Translate(
            fittedBounds.X,
            fittedBounds.Y,
            pageHeight,
            localPageHeight);
        PdfPagePatternSelection? localStrokePattern = invocation.StrokePattern?.Translate(
            fittedBounds.X,
            fittedBounds.Y,
            pageHeight,
            localPageHeight);
        groupTransform = OfficeTransform.Translate(fittedBounds.X, fittedBounds.Y);

        int failureVersion = type3GlyphBudget.FailureVersion;
        var elements = new List<PdfPageDrawingElement>();
        CollectVisualPrimitivesAndForms(
            content,
            resources,
            localTransform,
            localPageWidth,
            localPageHeight,
            primitive => elements.Add(PdfPageDrawingElement.FromPrimitive(primitive, elements.Count)),
            activeForms,
            activeType3Glyphs,
            renderedType3PaintOrders,
            type3GlyphBudget,
            invocation.PaintOrder,
            paintOrderScale * 0.000000001D,
            initialClipPath: localClipPath,
            initialFillColor: invocation.FillColor,
            initialFillColorSpace: invocation.FillColorSpace,
            initialFillPattern: localFillPattern,
            initialFillPatternBaseColorSpace: invocation.FillPatternBaseColorSpace,
            initialFillOpacity: 1D,
            initialStrokeColor: invocation.StrokeColor,
            initialStrokeColorSpace: invocation.StrokeColorSpace,
            initialStrokePattern: localStrokePattern,
            initialStrokePatternBaseColorSpace: invocation.StrokePatternBaseColorSpace,
            initialStrokeOpacity: 1D,
            initialStrokeWidth: invocation.StrokeWidth,
            initialStrokeDashStyle: invocation.StrokeDashStyle,
            initialStrokeLineCap: invocation.StrokeLineCap,
            initialStrokeLineJoin: invocation.StrokeLineJoin,
            contentNestingDepth: contentNestingDepth + 1,
            includeTilingPatterns: includeTilingPatterns,
            retainPrimitiveData: retainPrimitiveData,
            requireSupportedType3Content: true,
            allowSupportedType3Patterns: !requireNestedType3Uncolored,
            allowSupportedType3TransparencyGroups: true,
            requireNestedType3Uncolored: requireNestedType3Uncolored,
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
            return Type3TransparencyGroupDrawingResult.Unsupported;
        }

        if (requireNestedType3Uncolored) {
            for (int elementIndex = 0; elementIndex < elements.Count; elementIndex++) {
                PdfPageDrawingElement element = elements[elementIndex];
                if (element.Kind != PdfPageDrawingElementKind.Primitive) continue;
                if (!TryApplyInheritedType3PatternPaint(
                        element.Primitive,
                        invocation.FillColor,
                        localFillPattern,
                        invocation.StrokeColor,
                        localStrokePattern,
                        localPageHeight,
                        out PdfPageVisualPrimitive paintedPrimitive)) {
                    groupDrawing = null!;
                    return Type3TransparencyGroupDrawingResult.Unsupported;
                }
                elements[elementIndex] = PdfPageDrawingElement
                    .FromPrimitive(paintedPrimitive, element.Sequence)
                    .WithEffect(element.Effect);
            }
        }

        var placements = new List<PdfImagePlacement>();
        CollectImagePlacementsAndForms(
            content,
            resources,
            0,
            localTransform,
            localPageHeight,
            placements,
            activeForms,
            invocation.FillColor,
            invocation.FillColorSpace,
            1D,
            invocation.PaintOrder,
            paintOrderScale * 0.000000001D,
            initialClipPath: localClipPath,
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
                    return Type3TransparencyGroupDrawingResult.Unsupported;
                }
                if (requireNestedType3Uncolored && !image.IsImageMask) {
                    groupDrawing = null!;
                    return Type3TransparencyGroupDrawingResult.Unsupported;
                }
                if (requireNestedType3Uncolored && localFillPattern.HasValue) {
                    Type3PatternImageMaskDrawingResult result = TryCreateInheritedPatternImageMaskDrawing(
                            localFillPattern,
                            placements[i],
                            image,
                            localPageWidth,
                            localPageHeight,
                            out OfficeDrawing? maskedPattern,
                            out OfficeTransform maskedPatternTransform);
                    if (result == Type3PatternImageMaskDrawingResult.Unsupported) {
                        groupDrawing = null!;
                        return Type3TransparencyGroupDrawingResult.Unsupported;
                    }
                    if (result == Type3PatternImageMaskDrawingResult.Invisible) continue;
                    elements.Add(PdfPageDrawingElement.FromGroup(
                        maskedPattern,
                        maskedPatternTransform,
                        placements[i].PaintOrder,
                        placements[i].ContentOrderKey,
                        elements.Count));
                } else {
                    elements.Add(PdfPageDrawingElement.FromImage(placements[i], image, elements.Count));
                }
            }
        }

        var effects = new List<PdfPageDrawingEffectTransition>();
        CollectGraphicsEffectTransitions(
            content,
            resources,
            localTransform,
            localPageHeight,
            effects,
            new HashSet<PdfStream>(),
            PdfPageDrawingEffect.Default,
            invocation.PaintOrder,
            paintOrderScale * 0.000000001D,
            initialClipPath: localClipPath,
            initialFillColor: invocation.FillColor,
            initialFillColorSpace: invocation.FillColorSpace,
            initialFillOpacity: 1D,
            initialStrokeColor: invocation.StrokeColor,
            initialStrokeColorSpace: invocation.StrokeColorSpace,
            initialStrokeOpacity: 1D,
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

        var contentDrawing = new OfficeDrawing(localPageWidth, localPageHeight);
        var softMasks = new Dictionary<(PdfStream Group, OfficeSoftMaskMode Mode, OfficeColor Backdrop, Matrix2D Transform, double Width, double Height), OfficeDrawingSoftMask>();
        var activeSoftMasks = new HashSet<PdfStream>();
        TextContentParser.TextOutputBudget outputBudget = textOutputBudget ?? CreateTextOutputBudget();
        for (int i = 0; i < elements.Count; i++) {
            AddDrawingElement(contentDrawing, localPageHeight, localTransform, elements[i], softMasks, activeSoftMasks, outputBudget, pageContentBudget, type3GlyphBudget);
        }
        double groupOpacity = invocation.FillOpacity ?? 1D;
        if (groupOpacity < 1D) {
            groupDrawing = new OfficeDrawing(localPageWidth, localPageHeight);
            groupDrawing.AddEffectDrawing(contentDrawing, OfficeTransform.Identity, groupOpacity);
        } else {
            groupDrawing = contentDrawing;
        }
        return Type3TransparencyGroupDrawingResult.Success;
    }

    private Type3TransparencyGroupDrawingResult TryGetVisibleType3TransparencyGroupBounds(
        PdfDictionary formDictionary,
        Matrix2D transform,
        PdfPageClipPath? activeClip,
        double pageWidth,
        double pageHeight,
        out PdfPageClipPath bounds) {
        bounds = default;
        if (!TryReadBox(
                formDictionary.Items.TryGetValue("BBox", out PdfObject? bboxObject) ? bboxObject : null,
                out (double X1, double Y1, double X2, double Y2) bbox) ||
            bbox.X2 <= bbox.X1 ||
            bbox.Y2 <= bbox.Y1) {
            return Type3TransparencyGroupDrawingResult.Unsupported;
        }

        (double X, double Y) transformedTopLeft = transform.Transform(bbox.X1, bbox.Y1);
        (double X, double Y) transformedTopRight = transform.Transform(bbox.X2, bbox.Y1);
        (double X, double Y) transformedBottomLeft = transform.Transform(bbox.X1, bbox.Y2);
        (double X, double Y) transformedBottomRight = transform.Transform(bbox.X2, bbox.Y2);
        (double X, double Y) topLeft = (transformedTopLeft.X, pageHeight - transformedTopLeft.Y);
        (double X, double Y) topRight = (transformedTopRight.X, pageHeight - transformedTopRight.Y);
        (double X, double Y) bottomLeft = (transformedBottomLeft.X, pageHeight - transformedBottomLeft.Y);
        (double X, double Y) bottomRight = (transformedBottomRight.X, pageHeight - transformedBottomRight.Y);
        double left = Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomLeft.X, bottomRight.X));
        double top = Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomLeft.Y, bottomRight.Y));
        double right = Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomLeft.X, bottomRight.X));
        double bottom = Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomLeft.Y, bottomRight.Y));
        PdfPageClipPath projectedBounds;
        bool firstEdgeHorizontal = NearlyEqual(topLeft.Y, topRight.Y);
        bool firstEdgeVertical = NearlyEqual(topLeft.X, topRight.X);
        bool secondEdgeHorizontal = NearlyEqual(topRight.Y, bottomRight.Y);
        bool secondEdgeVertical = NearlyEqual(topRight.X, bottomRight.X);
        if ((firstEdgeHorizontal && secondEdgeVertical) || (firstEdgeVertical && secondEdgeHorizontal)) {
            projectedBounds = PdfPageClipPath.Rectangle(left, top, right - left, bottom - top);
        } else {
            var commands = new[] {
                OfficePathCommand.MoveTo(topLeft.X, topLeft.Y),
                OfficePathCommand.LineTo(topRight.X, topRight.Y),
                OfficePathCommand.LineTo(bottomRight.X, bottomRight.Y),
                OfficePathCommand.LineTo(bottomLeft.X, bottomLeft.Y),
                OfficePathCommand.Close()
            };
            if (!PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out projectedBounds)) {
                return Type3TransparencyGroupDrawingResult.Unsupported;
            }
        }
        if (activeClip.HasValue) {
            projectedBounds = PdfPageClipPath.ResolveActiveClip(projectedBounds, activeClip.Value);
        }
        if (projectedBounds.Width <= 0D || projectedBounds.Height <= 0D) {
            return Type3TransparencyGroupDrawingResult.Invisible;
        }
        if (!TryFitClipToDrawing(projectedBounds, pageWidth, pageHeight, out bounds)) {
            return Type3TransparencyGroupDrawingResult.Invisible;
        }
        if (bounds.IsRectangle) {
            const double localSurfaceTolerance = 0.000000001D;
            double leftWithTolerance = Math.Max(0D, bounds.X - localSurfaceTolerance);
            double topWithTolerance = Math.Max(0D, bounds.Y - localSurfaceTolerance);
            double rightWithTolerance = Math.Min(pageWidth, bounds.X + bounds.Width + localSurfaceTolerance);
            double bottomWithTolerance = Math.Min(pageHeight, bounds.Y + bounds.Height + localSurfaceTolerance);
            bounds = PdfPageClipPath.Rectangle(
                leftWithTolerance,
                topWithTolerance,
                rightWithTolerance - leftWithTolerance,
                bottomWithTolerance - topWithTolerance);
        }
        return Type3TransparencyGroupDrawingResult.Success;
    }

    private enum Type3TransparencyGroupDrawingResult {
        Success,
        Invisible,
        Unsupported
    }
}

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
        if (formDictionary.Items.TryGetValue("Matrix", out PdfObject? matrixObject)) {
            if (ResolveEffectObject(matrixObject) is not PdfArray matrix || matrix.Items.Count != 6) return false;
            for (int index = 0; index < matrix.Items.Count; index++) {
                if (ResolveEffectObject(matrix.Items[index]) is not PdfNumber number || !IsFinite(number.Value)) return false;
            }
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
        RenderedType3TextTracker renderedType3PaintOrders,
        Type3GlyphBudget type3GlyphBudget,
        double paintOrderScale,
        bool includeTilingPatterns,
        bool retainPrimitiveData,
        bool requireNestedType3Uncolored,
        TilingPatternResourceCache? tilingPatternResourceCache,
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

        PdfPageClipPath renderBounds = ExpandType3TransparencyGroupRenderSurface(fittedBounds, pageWidth, pageHeight);
        double localPageWidth = renderBounds.Width;
        double localPageHeight = renderBounds.Height;
        LocalizeType3TransparencyGroupProjection(
            transform,
            renderBounds,
            pageHeight,
            invocation.FillPattern,
            invocation.StrokePattern,
            out Matrix2D localTransform,
            out PdfPageClipPath? localClipPath,
            out PdfPagePatternSelection? localFillPattern,
            out PdfPagePatternSelection? localStrokePattern);
        groupTransform = OfficeTransform.Translate(fittedBounds.X, fittedBounds.Y);

        int failureVersion = type3GlyphBudget.FailureVersion;
        var elements = new List<PdfPageDrawingElement>();
        CollectVisualPrimitivesAndForms(
            content,
            resources,
            localTransform,
            localPageWidth,
            localPageHeight,
            primitive => {
                if ((primitive.ClipPath.HasValue && !primitive.ClipPath.Value.IsExact) ||
                    !CanRenderTilingPatterns(primitive, localPageWidth, localPageHeight)) {
                    type3GlyphBudget.RecordFailure();
                } else {
                    elements.Add(PdfPageDrawingElement.FromPrimitive(primitive, elements.Count));
                }
            },
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
            type3ImageVisitor: (placement, image, effect) => {
                if (!image.Interpolate) {
                    type3GlyphBudget.RecordFailure();
                } else {
                    elements.Add(PdfPageDrawingElement.FromImage(placement, image, elements.Count).WithEffect(effect));
                }
            },
            type3PrimitiveVisitor: (primitive, effect) => {
                if (!CanRenderTilingPatterns(primitive, localPageWidth, localPageHeight)) {
                    type3GlyphBudget.RecordFailure();
                } else {
                    elements.Add(PdfPageDrawingElement.FromPrimitive(primitive, elements.Count).WithEffect(effect));
                }
            },
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
                        localPageWidth,
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

        if (!localTransform.IsConformalStrokeTransform()) {
            for (int elementIndex = 0; elementIndex < elements.Count; elementIndex++) {
                PdfPageDrawingElement element = elements[elementIndex];
                if (element.Kind == PdfPageDrawingElementKind.Primitive &&
                    (ResolveVisibleType3PrimitivePaintChannels(element.Primitive, localPageWidth, localPageHeight) & PdfType3PaintChannels.Stroke) != 0) {
                    groupDrawing = null!;
                    return Type3TransparencyGroupDrawingResult.Unsupported;
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
            contentOrderPrefix: contentOrderPrefix,
            skipTransparencyGroupForms: true);
        SortGraphicsEffectTransitions(effects);

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
            var visiblePlacements = new List<PdfImagePlacement>(placements.Count);
            for (int i = 0; i < placements.Count; i++) {
                if (!IsInvisibleImagePlacement(placements[i], localPageHeight, localPageWidth, localPageHeight)) {
                    visiblePlacements.Add(placements[i]);
                }
            }
            var paintChannelCache = new Type3PaintChannelCache();
            var activePaintStreams = new HashSet<PdfStream>();
            for (int i = 0; i < visiblePlacements.Count; i++) {
                PdfImagePlacement placement = visiblePlacements[i];
                PdfPageDrawingEffect effect = ResolveDrawingEffect(
                    effects,
                    placement.PaintOrder,
                    contentOrderKey: placement.ContentOrderKey);
                bool suppressedBySoftMask = IsPaintSuppressedByTransparentSoftMask(
                        effect,
                        resources,
                        localTransform,
                        localPageWidth,
                        localPageHeight,
                        paintChannelCache,
                        activePaintStreams,
                        pageContentBudget,
                        type3GlyphBudget,
                        contentNestingDepth + 2);
                if (suppressedBySoftMask) {
                    continue;
                }
                PdfExtractedImage? image = GetImageForPlacement(resources, placement, colorizeImageMasks: true);
                if (image == null ||
                    !image.IsImageFile ||
                    !IsSupportedType3Image(placement, image, resources) ||
                    image.HasUnresolvedTransparencyMask ||
                    !TryCreateImageProjection(
                        placement,
                        localPageHeight,
                        localPageWidth,
                        localPageHeight,
                        out _,
                        allowAxisAlignedFallback: false)) {
                    groupDrawing = null!;
                    return Type3TransparencyGroupDrawingResult.Unsupported;
                }
                if (image.IsImageMask && placement.FillPattern.HasValue) {
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
                            placement,
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
                        placement.PaintOrder,
                        placement.ContentOrderKey,
                        elements.Count));
                } else {
                    elements.Add(PdfPageDrawingElement.FromImage(placement, image, elements.Count));
                }
            }
        }

        OverlayDrawingEffects(elements, effects);
        if (elements.Any(static element => element.Effect.BlendMode != OfficeBlendMode.Normal)) {
            groupDrawing = null!;
            return Type3TransparencyGroupDrawingResult.Unsupported;
        }
        SortDrawingElements(elements);

        var contentDrawing = new OfficeDrawing(localPageWidth, localPageHeight);
        var softMasks = new Dictionary<(PdfStream Group, PdfDictionary? ParentResources, OfficeSoftMaskMode Mode, OfficeColor Backdrop, Matrix2D Transform, double Width, double Height), OfficeDrawingSoftMask>();
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
        if (renderBounds.X != fittedBounds.X || renderBounds.Y != fittedBounds.Y ||
            renderBounds.Width != fittedBounds.Width || renderBounds.Height != fittedBounds.Height) {
            var exactDrawing = new OfficeDrawing(fittedBounds.Width, fittedBounds.Height);
            exactDrawing.AddDrawingForClippedRendering(
                groupDrawing,
                renderBounds.X - fittedBounds.X,
                renderBounds.Y - fittedBounds.Y,
                null);
            groupDrawing = exactDrawing;
        }
        return Type3TransparencyGroupDrawingResult.Success;
    }

    private static PdfPageClipPath ExpandType3TransparencyGroupRenderSurface(
        PdfPageClipPath bounds,
        double pageWidth,
        double pageHeight) {
        if (!bounds.IsRectangle) return bounds;
        const double localSurfaceTolerance = 0.000000001D;
        double left = Math.Max(0D, bounds.X - localSurfaceTolerance);
        double top = Math.Max(0D, bounds.Y - localSurfaceTolerance);
        double right = Math.Min(pageWidth, bounds.X + bounds.Width + localSurfaceTolerance);
        double bottom = Math.Min(pageHeight, bounds.Y + bounds.Height + localSurfaceTolerance);
        return PdfPageClipPath.Rectangle(left, top, right - left, bottom - top);
    }

    private static void LocalizeType3TransparencyGroupProjection(
        Matrix2D transform,
        PdfPageClipPath fittedBounds,
        double pageHeight,
        PdfPagePatternSelection? fillPattern,
        PdfPagePatternSelection? strokePattern,
        out Matrix2D localTransform,
        out PdfPageClipPath? localClipPath,
        out PdfPagePatternSelection? localFillPattern,
        out PdfPagePatternSelection? localStrokePattern) {
        double localPageHeight = fittedBounds.Height;
        localTransform = Matrix2D.Multiply(
            Matrix2D.Translation(
                -fittedBounds.X,
                localPageHeight - pageHeight + fittedBounds.Y),
            transform);
        localClipPath = fittedBounds.Translate(fittedBounds.X, fittedBounds.Y);
        localFillPattern = fillPattern?.Translate(
            fittedBounds.X,
            fittedBounds.Y,
            pageHeight,
            localPageHeight);
        localStrokePattern = strokePattern?.Translate(
            fittedBounds.X,
            fittedBounds.Y,
            pageHeight,
            localPageHeight);
    }

    private Type3TransparencyGroupDrawingResult TryGetVisibleType3TransparencyGroupBounds(
        PdfDictionary formDictionary,
        Matrix2D transform,
        PdfPageClipPath? activeClip,
        double pageWidth,
        double pageHeight,
        out PdfPageClipPath bounds) {
        bounds = default;
        if (!TryReadType3TransparencyGroupBox(
                formDictionary.Items.TryGetValue("BBox", out PdfObject? bboxObject) ? bboxObject : null,
                out (double X1, double Y1, double X2, double Y2) bbox)) {
            return Type3TransparencyGroupDrawingResult.Unsupported;
        }
        if (bbox.X2 <= bbox.X1 || bbox.Y2 <= bbox.Y1) {
            return Type3TransparencyGroupDrawingResult.Invisible;
        }

        (double X, double Y) transformedTopLeft = transform.Transform(bbox.X1, bbox.Y1);
        (double X, double Y) transformedTopRight = transform.Transform(bbox.X2, bbox.Y1);
        (double X, double Y) transformedBottomLeft = transform.Transform(bbox.X1, bbox.Y2);
        (double X, double Y) transformedBottomRight = transform.Transform(bbox.X2, bbox.Y2);
        if (!IsFinite(transformedTopLeft.X) || !IsFinite(transformedTopLeft.Y) ||
            !IsFinite(transformedTopRight.X) || !IsFinite(transformedTopRight.Y) ||
            !IsFinite(transformedBottomLeft.X) || !IsFinite(transformedBottomLeft.Y) ||
            !IsFinite(transformedBottomRight.X) || !IsFinite(transformedBottomRight.Y)) {
            return Type3TransparencyGroupDrawingResult.Unsupported;
        }
        (double X, double Y) topLeft = (transformedTopLeft.X, pageHeight - transformedTopLeft.Y);
        (double X, double Y) topRight = (transformedTopRight.X, pageHeight - transformedTopRight.Y);
        (double X, double Y) bottomLeft = (transformedBottomLeft.X, pageHeight - transformedBottomLeft.Y);
        (double X, double Y) bottomRight = (transformedBottomRight.X, pageHeight - transformedBottomRight.Y);
        double left = Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomLeft.X, bottomRight.X));
        double top = Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomLeft.Y, bottomRight.Y));
        double right = Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomLeft.X, bottomRight.X));
        double bottom = Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomLeft.Y, bottomRight.Y));
        if (!IsFinite(left) || !IsFinite(top) || !IsFinite(right) || !IsFinite(bottom) ||
            !IsFinite(right - left) || !IsFinite(bottom - top)) {
            return Type3TransparencyGroupDrawingResult.Unsupported;
        }
        PdfPageClipPath projectedBounds;
        bool firstEdgeHorizontal = topLeft.Y == topRight.Y;
        bool firstEdgeVertical = topLeft.X == topRight.X;
        bool secondEdgeHorizontal = topRight.Y == bottomRight.Y;
        bool secondEdgeVertical = topRight.X == bottomRight.X;
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
            if (projectedBounds.CanProveNoPositiveAreaIntersection(activeClip.Value)) {
                return Type3TransparencyGroupDrawingResult.Invisible;
            }
            projectedBounds = PdfPageClipPath.ResolveActiveClip(activeClip.Value, projectedBounds);
        }
        if (!projectedBounds.IsExact) {
            return Type3TransparencyGroupDrawingResult.Unsupported;
        }
        if (projectedBounds.Width <= 0D || projectedBounds.Height <= 0D) {
            return Type3TransparencyGroupDrawingResult.Invisible;
        }
        if (!TryFitClipToDrawing(projectedBounds, pageWidth, pageHeight, out bounds)) {
            return Type3TransparencyGroupDrawingResult.Invisible;
        }
        var geometryBudget = new VisualGeometryBudget();
        VisualPath? visibleBounds = VisualPath.FromClip(bounds, geometryBudget);
        if (visibleBounds != null && !geometryBudget.Exceeded &&
            !VisualPath.HasPositiveAreaIntersection(new[] { visibleBounds }, geometryBudget) &&
            !geometryBudget.Exceeded) {
            return Type3TransparencyGroupDrawingResult.Invisible;
        }
        return Type3TransparencyGroupDrawingResult.Success;
    }

    private bool TryReadType3TransparencyGroupBox(
        PdfObject? value,
        out (double X1, double Y1, double X2, double Y2) box) {
        box = default;
        PdfArray? array = ResolveArray(value);
        if (array == null || array.Items.Count != 4 ||
            ResolveObject(array.Items[0]) is not PdfNumber x1 ||
            ResolveObject(array.Items[1]) is not PdfNumber y1 ||
            ResolveObject(array.Items[2]) is not PdfNumber x2 ||
            ResolveObject(array.Items[3]) is not PdfNumber y2 ||
            !IsFinite(x1.Value) ||
            !IsFinite(y1.Value) ||
            !IsFinite(x2.Value) ||
            !IsFinite(y2.Value)) {
            return false;
        }

        box = (
            Math.Min(x1.Value, x2.Value),
            Math.Min(y1.Value, y2.Value),
            Math.Max(x1.Value, x2.Value),
            Math.Max(y1.Value, y2.Value));
        return true;
    }

    private enum Type3TransparencyGroupDrawingResult {
        Success,
        Invisible,
        Unsupported
    }
}

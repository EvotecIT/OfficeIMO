using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private bool RenderType3TextInvocation(
        PdfPageType3TextInvocation invocation,
        double pageWidth,
        double pageHeight,
        Action<PdfPageVisualPrimitive> primitiveVisitor,
        HashSet<PdfStream> activeForms,
        HashSet<PdfStream> activeType3Glyphs,
        Type3GlyphBudget type3GlyphBudget,
        double paintOrderScale,
        bool includeTilingPatterns,
        bool retainPrimitiveData,
        Dictionary<(PdfStream Stream, PdfDictionary Resources), PdfPageTilingPatternResource?>? tilingPatternResourceCache,
        TextContentParser.TextOutputBudget? textOutputBudget,
        PageContentBudget pageContentBudget,
        int contentNestingDepth,
        Action<PdfImagePlacement, PdfExtractedImage, PdfPageDrawingEffect>? imageVisitor,
        Action<PdfPageVisualPrimitive, PdfPageDrawingEffect>? primitiveEffectVisitor,
        Action<OfficeDrawing, OfficeTransform, double, PdfContentOrderKey?, PdfPageDrawingEffect>? groupVisitor,
        PdfContentOrderKey? contentOrderPrefix) {
        if (invocation.Glyphs.Count == 0) return false;

        for (int i = 0; i < invocation.Glyphs.Count; i++) {
            PdfPageType3GlyphInvocation glyph = invocation.Glyphs[i];
            if (glyph.Font.Type3 is not PdfType3FontResource type3 ||
                !type3.TryGetGlyph(glyph.CharacterCode, out PdfStream glyphStream) ||
                Filters.StreamDecoder.GetUnsupportedFilters(glyphStream.Dictionary, _objects).Count != 0 ||
                activeType3Glyphs.Contains(glyphStream)) return false;
        }

        var glyphPrimitives = new List<(PdfPageVisualPrimitive Primitive, PdfPageDrawingEffect Effect)>();
        var glyphImages = new List<(PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect)>();
        var glyphGroups = new List<(OfficeDrawing Drawing, OfficeTransform Transform, double PaintOrder, PdfContentOrderKey? ContentOrderKey, PdfPageDrawingEffect Effect)>();
        var extractedImageCache = new Dictionary<(int ObjectNumber, int DirectStreamIdentity, string ResourceName, OfficeColor MaskColor), PdfExtractedImage>();
        var validatedSoftMaskGroups = new HashSet<PdfStream>();
        var softMaskValidationBudget = new PageContentBudget(this);
        double nextPaintOrder = invocation.PaintOrder;
        double paintOrderLimit = invocation.PaintOrder + (Math.Abs(paintOrderScale) * 0.5D);
        for (int i = 0; i < invocation.Glyphs.Count; i++) {
            PdfPageType3GlyphInvocation glyph = invocation.Glyphs[i];
            PdfType3FontResource type3 = glyph.Font.Type3!;
            _ = type3.TryGetGlyph(glyph.CharacterCode, out PdfStream glyphStream);
            if (!activeType3Glyphs.Add(glyphStream)) return false;
            try {
                PdfContentOrderKey glyphOrderPrefix = (contentOrderPrefix ?? PdfContentOrderKey.Root).Append(i);
                var localPrimitives = new List<(PdfPageVisualPrimitive Primitive, PdfPageDrawingEffect Effect)>();
                var localImages = new List<(PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect)>();
                var localGroups = new List<(OfficeDrawing Drawing, OfficeTransform Transform, double PaintOrder, PdfContentOrderKey? ContentOrderKey, PdfPageDrawingEffect Effect)>();
                var localImagePlacements = new List<PdfImagePlacement>();
                int failureVersion = type3GlyphBudget.FailureVersion;
                string glyphContent;
                try {
                    glyphContent = PdfEncoding.Latin1GetString(pageContentBudget.Decode(glyphStream));
                } catch (IOException exception) when (exception is not PdfReadLimitException) {
                    return false;
                }

                Matrix2D glyphTransform = Matrix2D.Multiply(glyph.Transform, type3.FontMatrix);
                var localRenderedType3PaintOrders = new HashSet<double>();
                try {
                    CollectVisualPrimitivesAndForms(
                        glyphContent,
                        type3.Resources,
                        glyphTransform,
                        pageWidth,
                        pageHeight,
                        primitive => localPrimitives.Add((primitive, PdfPageDrawingEffect.Default)),
                        activeForms,
                        activeType3Glyphs,
                        localRenderedType3PaintOrders,
                        type3GlyphBudget,
                        0D,
                        1D,
                        initialClipPath: glyph.ClipPath,
                        initialFillColor: glyph.FillColor,
                        initialFillColorSpace: glyph.FillColorSpace,
                        initialFillPattern: glyph.FillPattern,
                        initialFillPatternBaseColorSpace: glyph.FillPatternBaseColorSpace,
                        initialFillOpacity: glyph.FillOpacity,
                        initialStrokeColor: glyph.StrokeColor,
                        initialStrokeColorSpace: glyph.StrokeColorSpace,
                        initialStrokePattern: glyph.StrokePattern,
                        initialStrokePatternBaseColorSpace: glyph.StrokePatternBaseColorSpace,
                        initialStrokeOpacity: glyph.StrokeOpacity,
                        initialStrokeWidth: glyph.StrokeWidth,
                        initialStrokeDashStyle: glyph.StrokeDashStyle,
                        initialStrokeLineCap: glyph.StrokeLineCap,
                        initialStrokeLineJoin: glyph.StrokeLineJoin,
                        contentNestingDepth: contentNestingDepth + 1,
                        includeTilingPatterns: includeTilingPatterns,
                        retainPrimitiveData: retainPrimitiveData,
                        requireSupportedType3Content: true,
                        allowSupportedType3Patterns: !type3.IsUncolored,
                        allowSupportedType3TransparencyGroups: true,
                        requireNestedType3Uncolored: type3.IsUncolored,
                        type3ImageVisitor: (placement, image, effect) => localImages.Add((placement, image, effect)),
                        type3PrimitiveVisitor: (primitive, effect) => localPrimitives.Add((primitive, effect)),
                        type3GroupVisitor: (drawing, transform, paintOrder, key, effect) => localGroups.Add((drawing, transform, paintOrder, key, effect)),
                        tilingPatternResourceCache: tilingPatternResourceCache,
                        textOutputBudget: textOutputBudget,
                        pageContentBudget: pageContentBudget,
                        contentOrderPrefix: glyphOrderPrefix);
                } catch (Exception exception) when (IsRecoverableType3ProjectionFailure(exception)) {
                    return false;
                }

                if (type3GlyphBudget.FailureVersion != failureVersion) return false;

                var localEffects = new List<PdfPageDrawingEffectTransition>();
                try {
                    CollectGraphicsEffectTransitions(
                        glyphContent,
                        type3.Resources,
                        glyphTransform,
                        pageHeight,
                        localEffects,
                        new HashSet<PdfStream>(),
                        PdfPageDrawingEffect.Default,
                        pageContentBudget: pageContentBudget,
                        contentOrderPrefix: glyphOrderPrefix,
                        skipTransparencyGroupForms: true);
                } catch (Exception exception) when (IsRecoverableType3ProjectionFailure(exception)) {
                    return false;
                }
                SortGraphicsEffectTransitions(localEffects);
                for (int effectIndex = 0; effectIndex < localEffects.Count; effectIndex++) {
                    if (!CanDecodeType3SoftMask(
                            localEffects[effectIndex].Effect.SoftMask,
                            softMaskValidationBudget,
                            validatedSoftMaskGroups)) {
                        return false;
                    }
                }
                for (int primitiveIndex = 0; primitiveIndex < localPrimitives.Count; primitiveIndex++) {
                    (PdfPageVisualPrimitive Primitive, PdfPageDrawingEffect Effect) item = localPrimitives[primitiveIndex];
                    PdfPageDrawingEffect inherited = ResolveDrawingEffect(localEffects, item.Primitive.PaintOrder, contentOrderKey: item.Primitive.ContentOrderKey);
                    localPrimitives[primitiveIndex] = (item.Primitive, item.Effect.OverlayOn(inherited));
                }
                for (int imageIndex = 0; imageIndex < localImages.Count; imageIndex++) {
                    (PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect) item = localImages[imageIndex];
                    PdfPageDrawingEffect inherited = ResolveDrawingEffect(localEffects, item.Placement.PaintOrder, contentOrderKey: item.Placement.ContentOrderKey);
                    localImages[imageIndex] = (item.Placement, item.Image, item.Effect.OverlayOn(inherited));
                }
                for (int groupIndex = 0; groupIndex < localGroups.Count; groupIndex++) {
                    (OfficeDrawing Drawing, OfficeTransform Transform, double PaintOrder, PdfContentOrderKey? ContentOrderKey, PdfPageDrawingEffect Effect) item = localGroups[groupIndex];
                    PdfPageDrawingEffect inherited = ResolveDrawingEffect(localEffects, item.PaintOrder, contentOrderKey: item.ContentOrderKey);
                    localGroups[groupIndex] = (item.Drawing, item.Transform, item.PaintOrder, item.ContentOrderKey, item.Effect.OverlayOn(inherited));
                }

                CollectImagePlacementsAndForms(
                    glyphContent,
                    type3.Resources,
                    0,
                    glyphTransform,
                    pageHeight,
                    localImagePlacements,
                    activeForms,
                    glyph.FillColor,
                    glyph.FillColorSpace,
                    glyph.FillOpacity,
                    0D,
                    1D,
                    initialClipPath: glyph.ClipPath,
                    contentNestingDepth: contentNestingDepth + 1,
                    pageContentBudget: pageContentBudget,
                    contentOrderPrefix: glyphOrderPrefix,
                    skipTransparencyGroupForms: true);
                if (type3.IsUncolored) {
                    for (int imageIndex = 0; imageIndex < localImages.Count; imageIndex++) {
                        (PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect) image = localImages[imageIndex];
                        if (!image.Image.IsImageMask) return false;
                        localImages[imageIndex] = (image.Placement.WithImageMaskColor(glyph.FillColor), image.Image, image.Effect);
                    }
                    for (int primitiveIndex = 0; primitiveIndex < localPrimitives.Count; primitiveIndex++) {
                        (PdfPageVisualPrimitive Primitive, PdfPageDrawingEffect Effect) item = localPrimitives[primitiveIndex];
                        if (!TryApplyInheritedType3PatternPaint(
                                item.Primitive,
                                glyph.FillColor,
                                glyph.FillPattern,
                                glyph.StrokeColor,
                                glyph.StrokePattern,
                                pageWidth,
                                pageHeight,
                                out PdfPageVisualPrimitive paintedPrimitive)) return false;
                        localPrimitives[primitiveIndex] = (paintedPrimitive, item.Effect);
                    }
                    for (int imageIndex = 0; imageIndex < localImagePlacements.Count; imageIndex++) {
                        localImagePlacements[imageIndex] = localImagePlacements[imageIndex].WithImageMaskColor(glyph.FillColor);
                    }
                }

                if (localImagePlacements.Count > 0) {
                    var pendingPlacements = new List<PdfImagePlacement>();
                    var pendingKeys = new HashSet<(int ObjectNumber, int DirectStreamIdentity, string ResourceName, OfficeColor MaskColor)>();
                    for (int imageIndex = 0; imageIndex < localImagePlacements.Count; imageIndex++) {
                        PdfImagePlacement placement = localImagePlacements[imageIndex];
                        var key = GetType3ImageCacheKey(placement);
                        if (!extractedImageCache.ContainsKey(key) && pendingKeys.Add(key)) pendingPlacements.Add(placement);
                    }
                    if (pendingPlacements.Count > 0) {
                        IReadOnlyList<PdfExtractedImage> images;
                        try {
                            images = GetImagesForResources(type3.Resources, 0, pendingPlacements, colorizeImageMasks: true);
                        } catch (IOException exception) when (exception is not PdfReadLimitException) {
                            return false;
                        } catch (NotSupportedException) {
                            return false;
                        }
                        for (int imageIndex = 0; imageIndex < pendingPlacements.Count; imageIndex++) {
                            PdfExtractedImage? image = FindImage(images, pendingPlacements[imageIndex]);
                            if (image == null || !image.IsImageFile) return false;
                            extractedImageCache[GetType3ImageCacheKey(pendingPlacements[imageIndex])] = image;
                        }
                    }
                    for (int imageIndex = 0; imageIndex < localImagePlacements.Count; imageIndex++) {
                        PdfExtractedImage image = extractedImageCache[GetType3ImageCacheKey(localImagePlacements[imageIndex])];
                        if (type3.IsUncolored && !image.IsImageMask) return false;
                        PdfImagePlacement placement = localImagePlacements[imageIndex];
                        PdfPageDrawingEffect effect = ResolveDrawingEffect(localEffects, placement.PaintOrder, contentOrderKey: placement.ContentOrderKey);
                        localImages.Add((placement, image, effect));
                    }
                }

                if (type3.IsUncolored && glyph.FillPattern.HasValue && localImages.Count > 0) {
                    for (int imageIndex = 0; imageIndex < localImages.Count; imageIndex++) {
                        (PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect) image = localImages[imageIndex];
                        Type3PatternImageMaskDrawingResult result = TryCreateInheritedPatternImageMaskDrawing(
                                glyph.FillPattern,
                                image.Placement,
                                image.Image,
                                pageWidth,
                                pageHeight,
                                out OfficeDrawing? maskedPattern,
                                out OfficeTransform maskedPatternTransform);
                        if (result == Type3PatternImageMaskDrawingResult.Unsupported) return false;
                        if (result == Type3PatternImageMaskDrawingResult.Invisible) continue;
                        localGroups.Add((maskedPattern, maskedPatternTransform, image.Placement.PaintOrder, image.Placement.ContentOrderKey, image.Effect));
                    }
                    localImages.Clear();
                }

                if (!TryPublishType3GlyphContent(
                        localPrimitives,
                        localImages,
                        localGroups,
                        ref nextPaintOrder,
                        paintOrderLimit,
                        glyphPrimitives,
                        glyphImages,
                        glyphGroups)) {
                    return false;
                }
            } finally {
                activeType3Glyphs.Remove(glyphStream);
            }
        }

        if (glyphImages.Count > 0 && imageVisitor == null) return false;
        if (glyphGroups.Count > 0 && groupVisitor == null) return false;
        for (int i = 0; i < glyphPrimitives.Count; i++) {
            if (primitiveEffectVisitor != null) primitiveEffectVisitor(glyphPrimitives[i].Primitive, glyphPrimitives[i].Effect);
            else primitiveVisitor(glyphPrimitives[i].Primitive);
        }
        for (int i = 0; i < glyphImages.Count; i++) imageVisitor!(glyphImages[i].Placement, glyphImages[i].Image, glyphImages[i].Effect);
        for (int i = 0; i < glyphGroups.Count; i++) {
            (OfficeDrawing Drawing, OfficeTransform Transform, double PaintOrder, PdfContentOrderKey? ContentOrderKey, PdfPageDrawingEffect Effect) group = glyphGroups[i];
            groupVisitor!(group.Drawing, group.Transform, group.PaintOrder, group.ContentOrderKey, group.Effect);
        }
        return true;
    }

    internal static PdfType3PaintChannels ResolveVisibleType3PrimitivePaintChannels(
        PdfPageVisualPrimitive primitive,
        double? pageWidth = null,
        double? pageHeight = null) {
        var budget = new VisualGeometryBudget();
        if (!HasFinitePrimitiveGeometry(primitive, budget)) return PdfType3PaintChannels.None;
        if (budget.Exceeded) return PdfType3PaintChannels.Both;

        IReadOnlyList<VisualPath>? visibleClips = null;
        if (pageWidth.HasValue && pageHeight.HasValue) {
            if (!IsFinite(pageWidth.Value) || !IsFinite(pageHeight.Value) ||
                pageWidth.Value <= 0D || pageHeight.Value <= 0D) {
                return PdfType3PaintChannels.None;
            }
            VisualPath? surface = VisualPath.Rectangle(
                0D,
                0D,
                pageWidth.Value,
                pageHeight.Value,
                OfficeTransform.Identity,
                budget);
            if (surface == null) return budget.Exceeded ? PdfType3PaintChannels.Both : PdfType3PaintChannels.None;
            visibleClips = new[] { surface };
        }
        if (primitive.ClipPath.HasValue) {
            PdfPageClipPath authoredClip = primitive.ClipPath.Value;
            if (!HasFiniteClipGeometry(authoredClip, budget) || authoredClip.Width <= 0D || authoredClip.Height <= 0D) {
                return PdfType3PaintChannels.None;
            }
            VisualPath? clipPath = VisualPath.FromClip(authoredClip, budget);
            if (clipPath == null) return budget.Exceeded ? PdfType3PaintChannels.Both : PdfType3PaintChannels.None;
            visibleClips = visibleClips == null ? new[] { clipPath } : AppendClip(visibleClips, clipPath);
            if (!VisualPath.HasPositiveAreaIntersection(visibleClips, budget)) {
                return budget.Exceeded ? PdfType3PaintChannels.Both : PdfType3PaintChannels.None;
            }
        }

        PdfType3PaintChannels channels = PdfType3PaintChannels.None;
        if (primitive.HasFillPaint && HasVisibleOpacity(primitive.FillOpacity)) {
            VisualPath? fillPath = VisualPath.FromFill(primitive, budget);
            if (fillPath == null) {
                if (budget.Exceeded) channels |= PdfType3PaintChannels.Fill;
            } else if (visibleClips == null || fillPath.IntersectsFills(visibleClips, budget)) {
                channels |= PdfType3PaintChannels.Fill;
            }
        }
        if (primitive.HasStrokePaint && primitive.StrokeWidth > 0D && HasVisibleOpacity(primitive.StrokeOpacity)) {
            VisualPath? strokePath = VisualPath.FromStroke(primitive, budget);
            if (strokePath == null) {
                if (budget.Exceeded) channels |= PdfType3PaintChannels.Stroke;
            } else if (visibleClips == null || strokePath.StrokeIntersectsFills(visibleClips, primitive.StrokeWidth / 2D, budget)) {
                channels |= PdfType3PaintChannels.Stroke;
            }
        }
        return budget.Exceeded ? PdfType3PaintChannels.Both : channels;
    }

    private static Type3PatternImageMaskDrawingResult TryCreateInheritedPatternImageMaskDrawing(
        PdfPagePatternSelection? selection,
        PdfImagePlacement placement,
        PdfExtractedImage image,
        double pageWidth,
        double pageHeight,
        out OfficeDrawing drawing,
        out OfficeTransform drawingTransform) {
        drawing = new OfficeDrawing(1D, 1D);
        drawingTransform = OfficeTransform.Identity;
        Type3PatternImageMaskDrawingResult preparation = TryPrepareInheritedPatternImageMaskDrawing(
            selection,
            placement,
            image,
            pageWidth,
            pageHeight,
            out OfficeImageProjection projection,
            out PdfPageClipPath fitted,
            out PdfPageTilingPatternPaint? tilingPaint,
            out OfficeLinearGradient? fillGradient,
            out OfficeRadialGradient? fillRadialGradient,
            out _);
        if (preparation != Type3PatternImageMaskDrawingResult.Success) return preparation;
        drawing = new OfficeDrawing(fitted.Width, fitted.Height);
        drawingTransform = OfficeTransform.Translate(fitted.X, fitted.Y);
        PdfPageTilingPatternPaint? localTilingPaint = tilingPaint == null
            ? null
            : new PdfPageTilingPatternPaint(
                tilingPaint.Resource,
                tilingPaint.Transform.Then(OfficeTransform.Translate(-fitted.X, -fitted.Y)),
                tilingPaint.Tint,
                tilingPaint.Opacity);
        var localPaintBounds = PdfPageVisualPrimitive.Rectangle(
            0D,
            0D,
            fitted.Width,
            fitted.Height,
            OfficeColor.Black,
            null,
            0D,
            OfficeStrokeDashStyle.Solid,
            null,
            null,
            null,
            null,
            null,
            placement.PaintOrder).WithPaints(
                OfficeColor.Black,
                localTilingPaint,
                fillGradient,
                fillRadialGradient,
                OfficeColor.Black,
                null,
                null,
                null);

        var patternDrawing = new OfficeDrawing(fitted.Width, fitted.Height);
        AddVisualPrimitive(patternDrawing, localPaintBounds);
        if (patternDrawing.Elements.Count == 0) return Type3PatternImageMaskDrawingResult.Unsupported;

        var maskDrawing = new OfficeDrawing(fitted.Width, fitted.Height);
        OfficeImageProjection localProjection = projection.Translate(-fitted.X, -fitted.Y);
        (double localLeft, double localTop, double localRight, double localBottom) = localProjection.GetDestinationBounds();
        bool projectionFitsDrawing = localLeft >= -0.000001D &&
            localTop >= -0.000001D &&
            localRight <= fitted.Width + 0.000001D &&
            localBottom <= fitted.Height + 0.000001D;
        if (fitted.IsRectangle && projectionFitsDrawing) {
            maskDrawing.AddImage(
                image.Bytes,
                image.MimeType,
                localProjection,
                opacity: placement.ImageOpacity ?? 1D);
        } else {
            OfficeClipPath? localClip = fitted.ToOfficeClipPath(fitted.X, fitted.Y);
            if (localClip == null) return Type3PatternImageMaskDrawingResult.Unsupported;
            maskDrawing.AddClippedImage(
                image.Bytes,
                image.MimeType,
                localProjection,
                0D,
                0D,
                localClip,
                opacity: placement.ImageOpacity ?? 1D);
        }
        if (maskDrawing.Elements.Count == 0) return Type3PatternImageMaskDrawingResult.Unsupported;

        drawing.AddEffectDrawing(
            patternDrawing,
            OfficeTransform.Identity,
            OfficeBlendMode.Normal,
            new OfficeDrawingSoftMask(maskDrawing));
        return Type3PatternImageMaskDrawingResult.Success;
    }

    private static Type3PatternImageMaskDrawingResult TryPrepareInheritedPatternImageMaskDrawing(
        PdfPagePatternSelection? selection,
        PdfImagePlacement placement,
        PdfExtractedImage image,
        double pageWidth,
        double pageHeight,
        out OfficeImageProjection projection,
        out PdfPageClipPath fitted,
        out PdfPageTilingPatternPaint? tilingPaint,
        out OfficeLinearGradient? fillGradient,
        out OfficeRadialGradient? fillRadialGradient,
        out bool shadingPreparationFailed) {
        projection = default;
        fitted = default;
        tilingPaint = null;
        fillGradient = null;
        fillRadialGradient = null;
        shadingPreparationFailed = false;
        if (!selection.HasValue || !image.IsImageMask || !image.IsImageFile) return Type3PatternImageMaskDrawingResult.Unsupported;
        if (!TryCreateImageProjection(
                placement,
                pageHeight,
                pageWidth,
                pageHeight,
                out projection,
                allowAxisAlignedFallback: false)) {
            return IsInvisibleImagePlacement(placement, pageHeight, pageWidth, pageHeight)
                ? Type3PatternImageMaskDrawingResult.Invisible
                : Type3PatternImageMaskDrawingResult.Unsupported;
        }

        (double left, double top, double right, double bottom) = projection.GetDestinationBounds();
        double width = right - left;
        double height = bottom - top;
        if (width <= 0D || height <= 0D) return Type3PatternImageMaskDrawingResult.Unsupported;

        PdfPageClipPath projectedBounds = PdfPageClipPath.Rectangle(left, top, width, height);
        if (placement.ClipPath.HasValue) {
            projectedBounds = PdfPageClipPath.ResolveActiveClip(projectedBounds, placement.ClipPath.Value);
        }
        if (!TryFitClipToDrawing(projectedBounds, pageWidth, pageHeight, out fitted)) {
            return Type3PatternImageMaskDrawingResult.Invisible;
        }
        if (!TryCreateInheritedTilingPatternPaint(selection, pageHeight, null, out tilingPaint)) {
            return Type3PatternImageMaskDrawingResult.Unsupported;
        }

        var globalPaintBounds = PdfPageVisualPrimitive.Rectangle(
            fitted.X,
            fitted.Y,
            fitted.Width,
            fitted.Height,
            OfficeColor.Black,
            null,
            0D,
            OfficeStrokeDashStyle.Solid,
            null,
            null,
            null,
            null,
            null,
            placement.PaintOrder);
        CreateInheritedShadingGradients(
            selection,
            globalPaintBounds,
            pageHeight,
            out fillGradient,
            out fillRadialGradient);
        if (selection.Value.ShadingPattern.HasValue && fillGradient == null && fillRadialGradient == null) {
            shadingPreparationFailed = true;
            return Type3PatternImageMaskDrawingResult.Unsupported;
        }
        return Type3PatternImageMaskDrawingResult.Success;
    }

    private enum Type3PatternImageMaskDrawingResult {
        Success,
        Invisible,
        Unsupported
    }

    private static bool TryCreateInheritedTilingPatternPaint(
        PdfPagePatternSelection? selection,
        double pageHeight,
        double? opacity,
        out PdfPageTilingPatternPaint? paint) {
        paint = null;
        if (!selection.HasValue) return true;
        if (selection.Value.ShadingPattern.HasValue) return true;
        PdfPageTilingPatternResource? resource = selection.Value.TilingPattern;
        if (resource == null) {
            return false;
        }
        if (resource.Uncolored && !selection.Value.Tint.HasValue) return false;

        var localToPattern = new Matrix2D(1D, 0D, 0D, -1D, resource.BoundingBoxX, resource.BoundingBoxTop);
        Matrix2D combined = Matrix2D.Multiply(
            new Matrix2D(1D, 0D, 0D, -1D, 0D, pageHeight),
            Matrix2D.Multiply(selection.Value.PaintTransform, Matrix2D.Multiply(resource.Matrix, localToPattern)));
        paint = new PdfPageTilingPatternPaint(
            resource,
            new OfficeTransform(combined.A, combined.B, combined.C, combined.D, combined.E, combined.F),
            resource.Uncolored ? selection.Value.Tint : null,
            opacity ?? 1D);
        return true;
    }

    private static bool IsUsableInheritedPattern(PdfPagePatternSelection? selection) {
        if (!selection.HasValue) return true;
        if (selection.Value.ShadingPattern.HasValue) return selection.Value.ShadingPattern.Value.SupportsExactType3Projection;
        PdfPageTilingPatternResource? pattern = selection.Value.TilingPattern;
        return pattern != null && (!pattern.Uncolored || selection.Value.Tint.HasValue);
    }

    private static bool TryApplyInheritedType3PatternPaint(
        PdfPageVisualPrimitive primitive,
        OfficeColor fillColor,
        PdfPagePatternSelection? fillPattern,
        OfficeColor strokeColor,
        PdfPagePatternSelection? strokePattern,
        double pageWidth,
        double pageHeight,
        out PdfPageVisualPrimitive paintedPrimitive) {
        paintedPrimitive = primitive;
        PdfType3PaintChannels visibleChannels = ResolveVisibleType3PrimitivePaintChannels(
            primitive,
            pageWidth,
            pageHeight);
        PdfPagePatternSelection? applicableFillPattern =
            (visibleChannels & PdfType3PaintChannels.Fill) != 0 ? fillPattern : null;
        PdfPagePatternSelection? applicableStrokePattern =
            (visibleChannels & PdfType3PaintChannels.Stroke) != 0 ? strokePattern : null;
        if (!IsUsableInheritedPattern(applicableFillPattern) ||
            !IsUsableInheritedPattern(applicableStrokePattern) ||
            !TryCreateInheritedTilingPatternPaint(
                applicableFillPattern,
                pageHeight,
                primitive.FillOpacity,
                out PdfPageTilingPatternPaint? fillPatternPaint) ||
            !TryCreateInheritedTilingPatternPaint(
                applicableStrokePattern,
                pageHeight,
                primitive.StrokeOpacity,
                out PdfPageTilingPatternPaint? strokePatternPaint)) {
            return false;
        }
        if (!CreateInheritedShadingGradients(
            applicableFillPattern,
            primitive,
            pageHeight,
            out OfficeLinearGradient? fillGradient,
            out OfficeRadialGradient? fillRadialGradient) ||
            !CreateInheritedShadingGradients(
            applicableStrokePattern,
            primitive,
            pageHeight,
            out OfficeLinearGradient? strokeGradient,
            out OfficeRadialGradient? strokeRadialGradient)) {
            return false;
        }
        if (applicableStrokePattern?.ShadingPattern.HasValue == true) {
            return false;
        }
        paintedPrimitive = primitive.WithPaints(
            fillColor,
            fillPatternPaint,
            fillGradient,
            fillRadialGradient,
            strokeColor,
            strokePatternPaint,
            strokeGradient,
            strokeRadialGradient);
        return true;
    }

    private static bool CreateInheritedShadingGradients(
        PdfPagePatternSelection? selection,
        PdfPageVisualPrimitive primitive,
        double pageHeight,
        out OfficeLinearGradient? linearGradient,
        out OfficeRadialGradient? radialGradient) {
        linearGradient = null;
        radialGradient = null;
        if (!selection.HasValue || !selection.Value.ShadingPattern.HasValue) return true;
        if (!PdfPageContentVisualParser.IsSupportedShadingTransform(
                selection.Value.ShadingPattern.Value,
                selection.Value.PaintTransform)) return false;
        Matrix2D combined = Matrix2D.Multiply(
            selection.Value.PaintTransform,
            selection.Value.ShadingPattern.Value.Matrix);
        if (!PdfPageContentVisualParser.IsSupportedExactShadingPlacement(
                selection.Value.ShadingPattern.Value.Shading,
                combined,
                primitive.X,
                primitive.Y,
                primitive.Width,
                primitive.Height,
                pageHeight)) return false;
        PdfPageContentVisualParser.CreateShadingGradients(
            selection.Value.ShadingPattern.Value,
            primitive.X,
            primitive.Y,
            primitive.Width,
            primitive.Height,
            selection.Value.PaintTransform,
            pageHeight,
            out linearGradient,
            out radialGradient);
        return linearGradient != null || radialGradient != null;
    }

    private static bool IsRecoverableType3ProjectionFailure(Exception exception) =>
        exception is not PdfReadLimitException &&
        (exception is IOException || exception is InvalidDataException || exception is NotSupportedException);

    private static (int ObjectNumber, int DirectStreamIdentity, string ResourceName, OfficeColor MaskColor) GetType3ImageCacheKey(PdfImagePlacement placement) =>
        (placement.ObjectNumber, placement.DirectStreamIdentity, placement.ResourceName, placement.ImageMaskColor);

    private static bool TryPublishType3GlyphContent(
        List<(PdfPageVisualPrimitive Primitive, PdfPageDrawingEffect Effect)> localPrimitives,
        List<(PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect)> localImages,
        List<(OfficeDrawing Drawing, OfficeTransform Transform, double PaintOrder, PdfContentOrderKey? ContentOrderKey, PdfPageDrawingEffect Effect)> localGroups,
        ref double nextPaintOrder,
        double paintOrderLimit,
        List<(PdfPageVisualPrimitive Primitive, PdfPageDrawingEffect Effect)> targetPrimitives,
        List<(PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect)> targetImages,
        List<(OfficeDrawing Drawing, OfficeTransform Transform, double PaintOrder, PdfContentOrderKey? ContentOrderKey, PdfPageDrawingEffect Effect)> targetGroups) {
        var items = new List<Type3GlyphPaintItem>(localPrimitives.Count + localImages.Count + localGroups.Count);
        for (int i = 0; i < localPrimitives.Count; i++) {
            items.Add(new Type3GlyphPaintItem(localPrimitives[i].Primitive.PaintOrder, localPrimitives[i].Primitive.ContentOrderKey, Type3GlyphPaintItemKind.Primitive, i, i));
        }
        for (int i = 0; i < localImages.Count; i++) {
            items.Add(new Type3GlyphPaintItem(localImages[i].Placement.PaintOrder, localImages[i].Placement.ContentOrderKey, Type3GlyphPaintItemKind.Image, i, localPrimitives.Count + i));
        }
        for (int i = 0; i < localGroups.Count; i++) {
            items.Add(new Type3GlyphPaintItem(localGroups[i].PaintOrder, localGroups[i].ContentOrderKey, Type3GlyphPaintItemKind.Group, i, localPrimitives.Count + localImages.Count + i));
        }
        items.Sort(static (left, right) => {
            if (left.ContentOrderKey != null && right.ContentOrderKey != null) {
                int contentOrder = left.ContentOrderKey.CompareTo(right.ContentOrderKey);
                if (contentOrder != 0) return contentOrder;
            }
            int order = left.PaintOrder.CompareTo(right.PaintOrder);
            return order != 0 ? order : left.Sequence.CompareTo(right.Sequence);
        });

        for (int i = 0; i < items.Count; i++) {
            nextPaintOrder = NextRepresentablePaintOrder(nextPaintOrder);
            if (nextPaintOrder >= paintOrderLimit) return false;
            Type3GlyphPaintItem item = items[i];
            if (item.Kind == Type3GlyphPaintItemKind.Image) {
                (PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect) image = localImages[item.Index];
                targetImages.Add((image.Placement.WithPaintOrder(nextPaintOrder), image.Image, image.Effect));
            } else if (item.Kind == Type3GlyphPaintItemKind.Group) {
                (OfficeDrawing Drawing, OfficeTransform Transform, double PaintOrder, PdfContentOrderKey? ContentOrderKey, PdfPageDrawingEffect Effect) group = localGroups[item.Index];
                targetGroups.Add((group.Drawing, group.Transform, nextPaintOrder, group.ContentOrderKey, group.Effect));
            } else {
                (PdfPageVisualPrimitive Primitive, PdfPageDrawingEffect Effect) primitive = localPrimitives[item.Index];
                targetPrimitives.Add((primitive.Primitive.WithPaintOrder(nextPaintOrder), primitive.Effect));
            }
        }
        return true;
    }
    private static double NextRepresentablePaintOrder(double value) {
        if (value == 0D) return double.Epsilon;
        long bits = BitConverter.DoubleToInt64Bits(value);
        bits += value > 0D ? 1L : -1L;
        return BitConverter.Int64BitsToDouble(bits);
    }

    private readonly struct Type3GlyphPaintItem {
        internal Type3GlyphPaintItem(double paintOrder, PdfContentOrderKey? contentOrderKey, Type3GlyphPaintItemKind kind, int index, int sequence) {
            PaintOrder = paintOrder;
            ContentOrderKey = contentOrderKey;
            Kind = kind;
            Index = index;
            Sequence = sequence;
        }

        internal double PaintOrder { get; }
        internal PdfContentOrderKey? ContentOrderKey { get; }
        internal Type3GlyphPaintItemKind Kind { get; }
        internal int Index { get; }
        internal int Sequence { get; }
    }

    private enum Type3GlyphPaintItemKind {
        Primitive,
        Image,
        Group
    }

    private PdfType3PaintChannels ResolveXObjectPaintChannels(
        PdfDictionary? resources,
        string name,
        PdfPageXObjectPaintState invocationState,
        double pageWidth,
        double pageHeight,
        Type3PaintChannelCache cache,
        HashSet<PdfStream> activeStreams,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        int depth = 0) {
        EnsureContentNestingBudget(depth);
        if (TryGetImageXObject(resources, name, out int objectNumber, out int directStreamIdentity)) {
            PdfImagePlacement placement = BuildImagePlacement(
                0,
                name,
                objectNumber,
                directStreamIdentity,
                invocationState.Transform,
                invocationState.ClipPath,
                OfficeColor.Black,
                invocationState.FillOpacity,
                paintOrder: 0D);
            return IsInvisibleImagePlacement(placement, pageHeight, pageWidth, pageHeight)
                ? PdfType3PaintChannels.None
                : PdfType3PaintChannels.Fill;
        }
        if (resources == null || !TryGetFormStream(resources, name, out PdfStream form)) {
            return PdfType3PaintChannels.Both;
        }
        if (form.Dictionary.Items.ContainsKey("Group")) {
            if (!HasVisibleOpacity(invocationState.FillOpacity)) {
                return PdfType3PaintChannels.None;
            }
            Type3TransparencyGroupDrawingResult boundsResult = TryGetVisibleType3TransparencyGroupBounds(
                form.Dictionary,
                ApplyFormMatrix(invocationState.Transform, form.Dictionary),
                invocationState.ClipPath,
                pageWidth,
                pageHeight,
                out _);
            if (boundsResult == Type3TransparencyGroupDrawingResult.Invisible) return PdfType3PaintChannels.None;
            if (boundsResult == Type3TransparencyGroupDrawingResult.Unsupported) return PdfType3PaintChannels.Both;
        }
        PdfDictionary formResources = ResolveDictionary(
            form.Dictionary.Items.TryGetValue("Resources", out PdfObject? value) ? value : null) ?? resources;
        return ResolveVisibleFormPaintChannels(
            form,
            formResources,
            invocationState,
            pageWidth,
            pageHeight,
            cache,
            activeStreams,
            pageContentBudget,
            type3GlyphBudget,
            depth);
    }

    private static bool IsInvisibleInlineImageInvocation(
        PdfPageXObjectInvocation invocation,
        PdfDictionary resources,
        double pageWidth,
        double pageHeight) {
        if (invocation.InlineImage == null) return false;
        PdfImagePlacement placement = BuildImagePlacement(
            0,
            invocation.InlineImage.ResourceName,
            0,
            invocation.InlineImage.DirectStreamIdentity,
            invocation.Transform,
            invocation.ClipPath,
            invocation.FillColor,
            invocation.FillOpacity,
            invocation.InlineImage.Stream,
            resources,
            invocation.PaintOrder);
        return IsInvisibleImagePlacement(placement, pageHeight, pageWidth, pageHeight);
    }

    private PdfType3PaintChannels ResolveType3PaintChannels(
        PdfStream stream,
        PdfDictionary resources,
        PdfPageXObjectPaintState programState,
        double pageWidth,
        double pageHeight,
        Type3PaintChannelCache cache,
        HashSet<PdfStream> activeStreams,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        EnsureContentNestingBudget(depth);
        var cacheKey = (
            Stream: stream,
            Resources: resources,
            ProgramState: programState,
            PageWidth: pageWidth,
            PageHeight: pageHeight);
        if (cache.Streams.TryGetValue(cacheKey, out PdfType3PaintChannels cached)) return cached;
        if (!activeStreams.Add(stream)) return PdfType3PaintChannels.Both;
        try {
            string content = PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream));
            PdfType3PaintChannels channels = PdfType3PaintChannels.None;
            Dictionary<string, PdfPageColorSpace> colorSpaces = GetColorSpaceResources(resources);
            Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
            Dictionary<string, Func<byte[], double>> widthProviders = ResourceResolver.GetFontWidthProvidersForResources(resources, _objects);
            IReadOnlyDictionary<string, PdfPageGraphicsStateResource> graphicsStates = GetGraphicsStateResources(resources);
            IReadOnlyList<PdfPageDrawingEffectTransition> effects = PdfPageGraphicsEffectTimelineParser.Parse(
                content,
                graphicsStates,
                PdfPageDrawingEffect.Default,
                programState.Transform,
                maxOperations: _limits.MaxContentOperations,
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands);
            string transformedContent = WrapContentWithTransform(content, programState.Transform, out int transformedOffset);
            _ = PdfPageContentVisualParser.Parse(
                transformedContent,
                pageWidth,
                pageHeight,
                graphicsStates,
                colorSpaces,
                GetShadingResources(resources),
                GetShadingPatternResources(resources),
                null,
                GetOptionalContentVisibility(resources),
                paintOrderOffset: -transformedOffset,
                initialClipPath: programState.ClipPath,
                initialFillOpacity: programState.FillOpacity,
                initialStrokeOpacity: programState.StrokeOpacity,
                initialStrokeWidth: programState.StrokeWidth,
                initialStrokeDashStyle: programState.StrokeDashStyle,
                initialStrokeLineCap: programState.StrokeLineCap,
                initialStrokeLineJoin: programState.StrokeLineJoin,
                maxOperations: _limits.MaxContentOperations,
                patternBaseColorSpaces: GetPatternBaseColorSpaceResources(resources),
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                primitiveVisitor: primitive => {
                    PdfPageDrawingEffect effect = ResolveDrawingEffect(effects, primitive.PaintOrder);
                    if (IsPaintSuppressedByTransparentSoftMask(
                            effect,
                            resources,
                            programState.Transform,
                            pageWidth,
                            pageHeight,
                            cache,
                            activeStreams,
                            pageContentBudget,
                            type3GlyphBudget,
                            depth + 1)) {
                        return;
                    }
                    channels |= ResolveVisibleType3PrimitivePaintChannels(primitive, pageWidth, pageHeight);
                },
                scaleStrokeWidthWithTransform: true,
                retainPrimitiveData: false);

            foreach (PdfPageXObjectInvocation invocation in PdfPageXObjectInvocationParser.Parse(
                         content,
                         programState.Transform,
                         pageHeight,
                         graphicsStates,
                         colorSpaces,
                         GetOptionalContentVisibility(resources),
                         initialFillOpacity: programState.FillOpacity,
                         initialClipPath: programState.ClipPath,
                         initialStrokeOpacity: programState.StrokeOpacity,
                         initialStrokeWidth: programState.StrokeWidth,
                         initialStrokeDashStyle: programState.StrokeDashStyle,
                         initialStrokeLineCap: programState.StrokeLineCap,
                         initialStrokeLineJoin: programState.StrokeLineJoin,
                         maxOperations: _limits.MaxContentOperations,
                         maxNestingDepth: _limits.MaxContentNestingDepth,
                         maxOperands: _limits.MaxContentOperands,
                         fonts: fonts,
                         fontWidthProviders: widthProviders,
                         type3TextVisitor: nested => {
                             for (int glyphIndex = 0; glyphIndex < nested.Glyphs.Count; glyphIndex++) {
                                 PdfPageType3GlyphInvocation glyph = nested.Glyphs[glyphIndex];
                                 channels |= ResolveType3PaintChannels(
                                     glyph,
                                     cache,
                                     activeStreams,
                                     pageContentBudget,
                                     type3GlyphBudget,
                                     pageWidth,
                                     pageHeight,
                                     depth + 1);
                             }
                             return true;
                         },
                         type3GlyphBudgetConsumer: type3GlyphBudget.Consume,
                         unsupportedTextVisitor: () => channels = PdfType3PaintChannels.Both,
                         xObjectPaintChannelResolver: (name, paintState) => ResolveXObjectPaintChannels(
                             resources,
                             name,
                             paintState,
                             pageWidth,
                             pageHeight,
                             cache,
                             activeStreams,
                             pageContentBudget,
                             type3GlyphBudget,
                             depth + 1),
                         softMaskVisibilityResolver: (softMask, transform) => !IsSoftMaskEntirelyTransparent(
                             softMask,
                             transform,
                             resources,
                             pageWidth,
                             pageHeight,
                             cache,
                             activeStreams,
                             pageContentBudget,
                             type3GlyphBudget,
                             depth + 1),
                         pageWidth: pageWidth)) {
                if (invocation.InlineImage != null) {
                    if (!IsInvisibleInlineImageInvocation(
                            invocation,
                            resources,
                            pageWidth,
                            pageHeight)) {
                        channels |= PdfType3PaintChannels.Fill;
                    }
                    continue;
                }
                channels |= ResolveXObjectPaintChannels(
                    resources,
                    invocation.Name,
                    invocation.PaintState,
                    pageWidth,
                    pageHeight,
                    cache,
                    activeStreams,
                    pageContentBudget,
                    type3GlyphBudget,
                    depth + 1);
            }
            cache.Streams[cacheKey] = channels;
            return channels;
        } catch (Exception exception) when (IsRecoverableType3ProjectionFailure(exception)) {
            return PdfType3PaintChannels.Both;
        } finally {
            activeStreams.Remove(stream);
        }
    }

    private sealed class Type3GlyphBudget {
        private readonly int _maximum;
        private int _count;
        private int _failureVersion;

        internal Type3GlyphBudget(int maximum) {
            _maximum = maximum;
        }

        internal void Consume(int count) {
            long next = (long)_count + count;
            if (next > _maximum) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.Type3GlyphInvocations, _maximum, next);
            }
            _count = (int)next;
        }

        internal int FailureVersion => _failureVersion;

        internal void RecordFailure() {
            _failureVersion++;
        }
    }
}

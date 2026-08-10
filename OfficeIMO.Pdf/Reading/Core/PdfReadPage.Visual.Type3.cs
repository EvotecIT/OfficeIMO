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
                        initialFillOpacity: glyph.FillOpacity,
                        initialStrokeColor: glyph.StrokeColor,
                        initialStrokeColorSpace: glyph.StrokeColorSpace,
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
                        requireNestedType3Uncolored: type3.IsUncolored,
                        type3ImageVisitor: (placement, image, effect) => localImages.Add((placement, image, effect)),
                        type3PrimitiveVisitor: (primitive, effect) => localPrimitives.Add((primitive, effect)),
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
                        contentOrderPrefix: glyphOrderPrefix);
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
                    contentOrderPrefix: glyphOrderPrefix);
                if (type3.IsUncolored) {
                    for (int imageIndex = 0; imageIndex < localImages.Count; imageIndex++) {
                        (PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect) image = localImages[imageIndex];
                        if (!image.Image.IsImageMask) return false;
                        localImages[imageIndex] = (image.Placement.WithImageMaskColor(glyph.FillColor), image.Image, image.Effect);
                    }
                    for (int primitiveIndex = 0; primitiveIndex < localPrimitives.Count; primitiveIndex++) {
                        (PdfPageVisualPrimitive Primitive, PdfPageDrawingEffect Effect) item = localPrimitives[primitiveIndex];
                        localPrimitives[primitiveIndex] = (item.Primitive.WithPaintColors(glyph.FillColor, glyph.StrokeColor), item.Effect);
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

                if (!TryPublishType3GlyphContent(
                        localPrimitives,
                        localImages,
                        ref nextPaintOrder,
                        paintOrderLimit,
                        glyphPrimitives,
                        glyphImages)) {
                    return false;
                }
            } finally {
                activeType3Glyphs.Remove(glyphStream);
            }
        }

        if (glyphImages.Count > 0 && imageVisitor == null) return false;
        for (int i = 0; i < glyphPrimitives.Count; i++) {
            if (primitiveEffectVisitor != null) primitiveEffectVisitor(glyphPrimitives[i].Primitive, glyphPrimitives[i].Effect);
            else primitiveVisitor(glyphPrimitives[i].Primitive);
        }
        for (int i = 0; i < glyphImages.Count; i++) imageVisitor!(glyphImages[i].Placement, glyphImages[i].Image, glyphImages[i].Effect);
        return true;
    }

    private static bool IsRecoverableType3ProjectionFailure(Exception exception) =>
        exception is not PdfReadLimitException &&
        (exception is IOException || exception is InvalidDataException || exception is NotSupportedException);

    private static (int ObjectNumber, int DirectStreamIdentity, string ResourceName, OfficeColor MaskColor) GetType3ImageCacheKey(PdfImagePlacement placement) =>
        (placement.ObjectNumber, placement.DirectStreamIdentity, placement.ResourceName, placement.ImageMaskColor);

    private static bool TryPublishType3GlyphContent(
        List<(PdfPageVisualPrimitive Primitive, PdfPageDrawingEffect Effect)> localPrimitives,
        List<(PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect)> localImages,
        ref double nextPaintOrder,
        double paintOrderLimit,
        List<(PdfPageVisualPrimitive Primitive, PdfPageDrawingEffect Effect)> targetPrimitives,
        List<(PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect)> targetImages) {
        var items = new List<Type3GlyphPaintItem>(localPrimitives.Count + localImages.Count);
        for (int i = 0; i < localPrimitives.Count; i++) {
            items.Add(new Type3GlyphPaintItem(localPrimitives[i].Primitive.PaintOrder, localPrimitives[i].Primitive.ContentOrderKey, false, i, i));
        }
        for (int i = 0; i < localImages.Count; i++) {
            items.Add(new Type3GlyphPaintItem(localImages[i].Placement.PaintOrder, localImages[i].Placement.ContentOrderKey, true, i, localPrimitives.Count + i));
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
            if (item.IsImage) {
                (PdfImagePlacement Placement, PdfExtractedImage Image, PdfPageDrawingEffect Effect) image = localImages[item.Index];
                targetImages.Add((image.Placement.WithPaintOrder(nextPaintOrder), image.Image, image.Effect));
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
        internal Type3GlyphPaintItem(double paintOrder, PdfContentOrderKey? contentOrderKey, bool isImage, int index, int sequence) {
            PaintOrder = paintOrder;
            ContentOrderKey = contentOrderKey;
            IsImage = isImage;
            Index = index;
            Sequence = sequence;
        }

        internal double PaintOrder { get; }
        internal PdfContentOrderKey? ContentOrderKey { get; }
        internal bool IsImage { get; }
        internal int Index { get; }
        internal int Sequence { get; }
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

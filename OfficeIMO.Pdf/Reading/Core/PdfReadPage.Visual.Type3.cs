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
        Action<PdfImagePlacement, PdfExtractedImage>? imageVisitor) {
        if (invocation.Glyphs.Count == 0) return false;

        for (int i = 0; i < invocation.Glyphs.Count; i++) {
            PdfPageType3GlyphInvocation glyph = invocation.Glyphs[i];
            if (glyph.Font.Type3 is not PdfType3FontResource type3 ||
                !type3.TryGetGlyph(glyph.CharacterCode, out PdfStream glyphStream) ||
                Filters.StreamDecoder.GetUnsupportedFilters(glyphStream.Dictionary, _objects).Count != 0 ||
                activeType3Glyphs.Contains(glyphStream)) return false;
        }

        var glyphPrimitives = new List<PdfPageVisualPrimitive>();
        var glyphImages = new List<(PdfImagePlacement Placement, PdfExtractedImage Image)>();
        var extractedImageCache = new Dictionary<(int ObjectNumber, int DirectStreamIdentity, string ResourceName, OfficeColor MaskColor), PdfExtractedImage>();
        double nextPaintOrder = invocation.PaintOrder;
        double paintOrderLimit = invocation.PaintOrder + (Math.Abs(paintOrderScale) * 0.5D);
        for (int i = 0; i < invocation.Glyphs.Count; i++) {
            PdfPageType3GlyphInvocation glyph = invocation.Glyphs[i];
            PdfType3FontResource type3 = glyph.Font.Type3!;
            _ = type3.TryGetGlyph(glyph.CharacterCode, out PdfStream glyphStream);
            if (!activeType3Glyphs.Add(glyphStream)) return false;
            try {
                var localPrimitives = new List<PdfPageVisualPrimitive>();
                var localImages = new List<(PdfImagePlacement Placement, PdfExtractedImage Image)>();
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
                        localPrimitives.Add,
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
                        type3ImageVisitor: (placement, image) => localImages.Add((placement, image)),
                        tilingPatternResourceCache: tilingPatternResourceCache,
                        textOutputBudget: textOutputBudget,
                        pageContentBudget: pageContentBudget);
                } catch (Exception exception) when (IsRecoverableType3ProjectionFailure(exception)) {
                    return false;
                }

                if (type3GlyphBudget.FailureVersion != failureVersion) return false;

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
                    pageContentBudget: pageContentBudget);
                if (type3.IsUncolored) {
                    if (localImages.Count > 0) return false;
                    for (int primitiveIndex = 0; primitiveIndex < localPrimitives.Count; primitiveIndex++) {
                        localPrimitives[primitiveIndex] = localPrimitives[primitiveIndex].WithPaintColors(glyph.FillColor, glyph.StrokeColor);
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
                        localImages.Add((localImagePlacements[imageIndex], image));
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
        for (int i = 0; i < glyphPrimitives.Count; i++) primitiveVisitor(glyphPrimitives[i]);
        for (int i = 0; i < glyphImages.Count; i++) imageVisitor!(glyphImages[i].Placement, glyphImages[i].Image);
        return true;
    }

    private static bool IsRecoverableType3ProjectionFailure(Exception exception) =>
        exception is not PdfReadLimitException &&
        (exception is IOException || exception is InvalidDataException || exception is NotSupportedException);

    private static (int ObjectNumber, int DirectStreamIdentity, string ResourceName, OfficeColor MaskColor) GetType3ImageCacheKey(PdfImagePlacement placement) =>
        (placement.ObjectNumber, placement.DirectStreamIdentity, placement.ResourceName, placement.ImageMaskColor);

    private static bool TryPublishType3GlyphContent(
        List<PdfPageVisualPrimitive> localPrimitives,
        List<(PdfImagePlacement Placement, PdfExtractedImage Image)> localImages,
        ref double nextPaintOrder,
        double paintOrderLimit,
        List<PdfPageVisualPrimitive> targetPrimitives,
        List<(PdfImagePlacement Placement, PdfExtractedImage Image)> targetImages) {
        var primitiveOrders = new HashSet<double>();
        for (int i = 0; i < localPrimitives.Count; i++) primitiveOrders.Add(localPrimitives[i].PaintOrder);
        for (int i = 0; i < localImages.Count; i++) {
            if (primitiveOrders.Contains(localImages[i].Placement.PaintOrder)) return false;
        }

        var items = new List<Type3GlyphPaintItem>(localPrimitives.Count + localImages.Count);
        for (int i = 0; i < localPrimitives.Count; i++) {
            items.Add(new Type3GlyphPaintItem(localPrimitives[i].PaintOrder, false, i, i));
        }
        for (int i = 0; i < localImages.Count; i++) {
            items.Add(new Type3GlyphPaintItem(localImages[i].Placement.PaintOrder, true, i, localPrimitives.Count + i));
        }
        items.Sort(static (left, right) => {
            int order = left.PaintOrder.CompareTo(right.PaintOrder);
            return order != 0 ? order : left.Sequence.CompareTo(right.Sequence);
        });

        for (int i = 0; i < items.Count; i++) {
            nextPaintOrder = NextRepresentablePaintOrder(nextPaintOrder);
            if (nextPaintOrder >= paintOrderLimit) return false;
            Type3GlyphPaintItem item = items[i];
            if (item.IsImage) {
                (PdfImagePlacement Placement, PdfExtractedImage Image) image = localImages[item.Index];
                targetImages.Add((image.Placement.WithPaintOrder(nextPaintOrder), image.Image));
            } else {
                targetPrimitives.Add(localPrimitives[item.Index].WithPaintOrder(nextPaintOrder));
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
        internal Type3GlyphPaintItem(double paintOrder, bool isImage, int index, int sequence) {
            PaintOrder = paintOrder;
            IsImage = isImage;
            Index = index;
            Sequence = sequence;
        }

        internal double PaintOrder { get; }
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

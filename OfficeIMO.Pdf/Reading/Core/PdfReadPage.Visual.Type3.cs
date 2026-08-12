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
        int contentNestingDepth) {
        if (invocation.Glyphs.Count == 0) return false;

        for (int i = 0; i < invocation.Glyphs.Count; i++) {
            PdfPageType3GlyphInvocation glyph = invocation.Glyphs[i];
            if (glyph.Font.Type3 is not PdfType3FontResource type3 ||
                !type3.TryGetGlyph(glyph.CharacterCode, out PdfStream glyphStream) ||
                Filters.StreamDecoder.GetUnsupportedFilters(glyphStream.Dictionary, _objects).Count != 0 ||
                activeType3Glyphs.Contains(glyphStream)) return false;
        }

        var glyphPrimitives = new List<PdfPageVisualPrimitive>();
        double nextPaintOrder = invocation.PaintOrder;
        double paintOrderLimit = invocation.PaintOrder + (Math.Abs(paintOrderScale) * 0.5D);
        for (int i = 0; i < invocation.Glyphs.Count; i++) {
            PdfPageType3GlyphInvocation glyph = invocation.Glyphs[i];
            PdfType3FontResource type3 = glyph.Font.Type3!;
            _ = type3.TryGetGlyph(glyph.CharacterCode, out PdfStream glyphStream);
            if (!activeType3Glyphs.Add(glyphStream)) return false;
            try {
                var localPrimitives = new List<PdfPageVisualPrimitive>();
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
                        contentNestingDepth: contentNestingDepth,
                        includeTilingPatterns: includeTilingPatterns,
                        retainPrimitiveData: retainPrimitiveData,
                        requireVectorOnly: true,
                        tilingPatternResourceCache: tilingPatternResourceCache,
                        textOutputBudget: textOutputBudget,
                        pageContentBudget: pageContentBudget);
                } catch (Exception exception) when (IsRecoverableType3ProjectionFailure(exception)) {
                    return false;
                }

                if (type3GlyphBudget.FailureVersion != failureVersion) return false;

                if (type3.IsUncolored) {
                    for (int primitiveIndex = 0; primitiveIndex < localPrimitives.Count; primitiveIndex++) {
                        localPrimitives[primitiveIndex] = localPrimitives[primitiveIndex].WithPaintColors(glyph.FillColor, glyph.StrokeColor);
                    }
                }

                localPrimitives.Sort(static (left, right) => left.PaintOrder.CompareTo(right.PaintOrder));
                for (int primitiveIndex = 0; primitiveIndex < localPrimitives.Count; primitiveIndex++) {
                    nextPaintOrder = NextRepresentablePaintOrder(nextPaintOrder);
                    if (nextPaintOrder >= paintOrderLimit) return false;
                    glyphPrimitives.Add(localPrimitives[primitiveIndex].WithPaintOrder(nextPaintOrder));
                }
            } finally {
                activeType3Glyphs.Remove(glyphStream);
            }
        }

        for (int i = 0; i < glyphPrimitives.Count; i++) primitiveVisitor(glyphPrimitives[i]);
        return true;
    }

    private static bool IsRecoverableType3ProjectionFailure(Exception exception) =>
        exception is not PdfReadLimitException &&
        (exception is IOException || exception is InvalidDataException || exception is NotSupportedException);

    private static double NextRepresentablePaintOrder(double value) {
        if (value == 0D) return double.Epsilon;
        long bits = BitConverter.DoubleToInt64Bits(value);
        bits += value > 0D ? 1L : -1L;
        return BitConverter.Int64BitsToDouble(bits);
    }

    private static bool TryPublishType3LocalPrimitives(
        List<PdfPageVisualPrimitive> localPrimitives,
        double paintOrder,
        double paintOrderScale,
        Action<PdfPageVisualPrimitive> primitiveVisitor) {
        if (localPrimitives.Count == 0) return true;

        var ordered = new List<PdfPageVisualPrimitive>(localPrimitives);
        ordered.Sort(static (left, right) => left.PaintOrder.CompareTo(right.PaintOrder));
        double nextPaintOrder = paintOrder;
        double paintOrderLimit = paintOrder + (Math.Abs(paintOrderScale) * 0.5D);
        for (int index = 0; index < ordered.Count; index++) {
            nextPaintOrder = NextRepresentablePaintOrder(nextPaintOrder);
            if (nextPaintOrder >= paintOrderLimit) return false;
            primitiveVisitor(ordered[index].WithPaintOrder(nextPaintOrder));
        }

        return true;
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

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
        HashSet<double> renderedType3PaintOrders,
        Type3GlyphBudget type3GlyphBudget,
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
        for (int i = 0; i < invocation.Glyphs.Count; i++) {
            PdfPageType3GlyphInvocation glyph = invocation.Glyphs[i];
            PdfType3FontResource type3 = glyph.Font.Type3!;
            _ = type3.TryGetGlyph(glyph.CharacterCode, out PdfStream glyphStream);
            if (!activeType3Glyphs.Add(glyphStream)) return false;
            try {
                int failureVersion = type3GlyphBudget.FailureVersion;
                string glyphContent;
                try {
                    glyphContent = PdfEncoding.Latin1GetString(pageContentBudget.Decode(glyphStream));
                } catch (IOException exception) when (exception is not PdfReadLimitException) {
                    return false;
                }

                Matrix2D glyphTransform = Matrix2D.Multiply(glyph.Transform, type3.FontMatrix);
                int primitiveStart = glyphPrimitives.Count;
                CollectVisualPrimitivesAndForms(
                    glyphContent,
                    type3.Resources,
                    glyphTransform,
                    pageWidth,
                    pageHeight,
                    glyphPrimitives.Add,
                    activeForms,
                    activeType3Glyphs,
                    renderedType3PaintOrders,
                    type3GlyphBudget,
                    invocation.PaintOrder + (i * 0.000000001D),
                    0.000000000001D,
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
                    requireVectorOnly: true,
                    tilingPatternResourceCache: tilingPatternResourceCache,
                    textOutputBudget: textOutputBudget,
                    pageContentBudget: pageContentBudget);

                if (type3GlyphBudget.FailureVersion != failureVersion) return false;

                if (type3.IsUncolored) {
                    for (int primitiveIndex = primitiveStart; primitiveIndex < glyphPrimitives.Count; primitiveIndex++) {
                        glyphPrimitives[primitiveIndex] = glyphPrimitives[primitiveIndex].WithPaintColors(glyph.FillColor, glyph.StrokeColor);
                    }
                }
            } finally {
                activeType3Glyphs.Remove(glyphStream);
            }
        }

        for (int i = 0; i < glyphPrimitives.Count; i++) primitiveVisitor(glyphPrimitives[i]);
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

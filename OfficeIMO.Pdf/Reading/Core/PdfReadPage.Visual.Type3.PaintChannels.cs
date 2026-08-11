namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private PdfType3PaintChannels ResolveVisibleFormPaintChannels(
        PdfStream form,
        PdfDictionary resources,
        Matrix2D invocationTransform,
        PdfPageClipPath? invocationClipPath,
        double pageWidth,
        double pageHeight,
        Dictionary<PdfStream, PdfType3PaintChannels> cache,
        HashSet<PdfStream> activeStreams,
        int depth) {
        if (!activeStreams.Add(form)) return PdfType3PaintChannels.Both;
        try {
            string content = WrapFormContentWithBoundingBoxClip(
                PdfEncoding.Latin1GetString(new PageContentBudget(this).Decode(form)),
                form.Dictionary);
            Matrix2D formTransform = ApplyFormMatrix(invocationTransform, form.Dictionary);
            PdfType3PaintChannels channels = PdfType3PaintChannels.None;
            Dictionary<string, PdfPageColorSpace> colorSpaces = GetColorSpaceResources(resources);
            var geometryBudget = new VisualGeometryBudget();
            var patternPaintCache = new Dictionary<PdfPageTilingPatternResource, bool>();
            _ = PdfPageContentVisualParser.Parse(
                WrapContentWithTransform(content, formTransform),
                pageWidth,
                pageHeight,
                GetGraphicsStateResources(resources),
                colorSpaces,
                GetShadingResources(resources),
                GetShadingPatternResources(resources),
                null,
                GetOptionalContentVisibility(resources),
                initialClipPath: invocationClipPath,
                maxOperations: _limits.MaxContentOperations,
                patternBaseColorSpaces: GetPatternBaseColorSpaceResources(resources),
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                primitiveVisitor: primitive => {
                    if (!IsVisibleVisualPrimitive(primitive, pageWidth, pageHeight, geometryBudget, patternPaintCache)) return;
                    if (primitive.HasFillPaint) channels |= PdfType3PaintChannels.Fill;
                    if (primitive.HasStrokePaint) channels |= PdfType3PaintChannels.Stroke;
                },
                retainPrimitiveData: false);

            Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
            Dictionary<string, Func<byte[], double>> widthProviders = ResourceResolver.GetFontWidthProvidersForResources(resources, _objects);
            foreach (PdfPageXObjectInvocation invocation in PdfPageXObjectInvocationParser.Parse(
                         content,
                         formTransform,
                         pageHeight,
                         GetGraphicsStateResources(resources),
                         colorSpaces,
                         GetOptionalContentVisibility(resources),
                         initialClipPath: invocationClipPath,
                         maxOperations: _limits.MaxContentOperations,
                         maxNestingDepth: _limits.MaxContentNestingDepth,
                         maxOperands: _limits.MaxContentOperands,
                         fonts: fonts,
                         fontWidthProviders: widthProviders,
                         type3TextVisitor: nested => {
                             for (int glyphIndex = 0; glyphIndex < nested.Glyphs.Count; glyphIndex++) {
                                 PdfPageType3GlyphInvocation glyph = nested.Glyphs[glyphIndex];
                                 if (glyph.Font.Type3 is not PdfType3FontResource nestedType3 ||
                                     !nestedType3.TryGetGlyph(glyph.CharacterCode, out PdfStream nestedStream)) {
                                     channels = PdfType3PaintChannels.Both;
                                     return true;
                                 }
                                 channels |= ResolveType3PaintChannels(
                                     nestedStream,
                                     nestedType3.Resources,
                                     cache,
                                     activeStreams,
                                     depth + 1);
                             }
                             return true;
                         },
                         unsupportedTextVisitor: () => channels = PdfType3PaintChannels.Both,
                         type3PaintChannelResolver: (font, bytes) => ResolveType3PaintChannels(font, bytes, cache, activeStreams),
                         xObjectPaintChannelResolver: (name, transform, clipPath) => ResolveXObjectPaintChannels(
                             resources,
                             name,
                             transform,
                             clipPath,
                             pageWidth,
                             pageHeight,
                             cache,
                             activeStreams,
                             depth + 1))) {
                if (invocation.InlineImage != null || TryGetImageXObject(resources, invocation.Name, out _, out _)) {
                    if ((invocation.FillOpacity ?? 1D) > 0D) channels |= PdfType3PaintChannels.Fill;
                    continue;
                }
                channels |= ResolveXObjectPaintChannels(
                    resources,
                    invocation.Name,
                    invocation.Transform,
                    invocation.ClipPath,
                    pageWidth,
                    pageHeight,
                    cache,
                    activeStreams,
                    depth + 1);
            }
            return channels;
        } catch (Exception exception) when (IsRecoverableType3ProjectionFailure(exception)) {
            return PdfType3PaintChannels.Both;
        } finally {
            activeStreams.Remove(form);
        }
    }
}

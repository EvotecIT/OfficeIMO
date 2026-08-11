namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private PdfType3PaintChannels ResolveVisibleFormPaintChannels(
        PdfStream form,
        PdfDictionary resources,
        Matrix2D invocationTransform,
        PdfPageClipPath? invocationClipPath,
        double? fillOpacity,
        double? strokeOpacity,
        double pageWidth,
        double pageHeight,
        Type3PaintChannelCache cache,
        HashSet<PdfStream> activeStreams,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        EnsureContentNestingBudget(depth);
        var cacheKey = (
            Stream: form,
            Resources: resources,
            InvocationTransform: invocationTransform,
            InvocationClipPath: invocationClipPath,
            FillOpacity: fillOpacity,
            StrokeOpacity: strokeOpacity,
            PageWidth: pageWidth,
            PageHeight: pageHeight);
        if (cache.VisibleForms.TryGetValue(cacheKey, out PdfType3PaintChannels cached)) return cached;
        if (!activeStreams.Add(form)) return PdfType3PaintChannels.Both;
        try {
            string content = WrapFormContentWithBoundingBoxClip(
                PdfEncoding.Latin1GetString(pageContentBudget.Decode(form)),
                form.Dictionary);
            Matrix2D formTransform = ApplyFormMatrix(invocationTransform, form.Dictionary);
            PdfType3PaintChannels channels = PdfType3PaintChannels.None;
            Dictionary<string, PdfPageColorSpace> colorSpaces = GetColorSpaceResources(resources);
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
                initialFillOpacity: fillOpacity,
                initialStrokeOpacity: strokeOpacity,
                maxOperations: _limits.MaxContentOperations,
                patternBaseColorSpaces: GetPatternBaseColorSpaceResources(resources),
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                primitiveVisitor: primitive => {
                    channels |= ResolveVisibleType3PrimitivePaintChannels(primitive, pageWidth, pageHeight);
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
                         initialFillOpacity: fillOpacity,
                         initialStrokeOpacity: strokeOpacity,
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
                                     depth + 1);
                             }
                             return true;
                         },
                         type3GlyphBudgetConsumer: type3GlyphBudget.Consume,
                         unsupportedTextVisitor: () => channels = PdfType3PaintChannels.Both,
                         type3PaintChannelResolver: glyph => ResolveType3PaintChannels(glyph, cache, activeStreams, pageContentBudget, type3GlyphBudget, depth + 1),
                         xObjectPaintChannelResolver: (name, transform, clipPath, resolvedFillOpacity, resolvedStrokeOpacity) => ResolveXObjectPaintChannels(
                             resources,
                             name,
                             transform,
                             clipPath,
                             resolvedFillOpacity,
                             resolvedStrokeOpacity,
                             pageWidth,
                             pageHeight,
                             cache,
                             activeStreams,
                             pageContentBudget,
                             type3GlyphBudget,
                             depth + 1),
                         pageWidth: pageWidth)) {
                if (invocation.InlineImage != null &&
                    !IsInvisibleInlineImageInvocation(invocation, resources, pageWidth, pageHeight)) {
                    channels |= PdfType3PaintChannels.Fill;
                    continue;
                }
                channels |= ResolveXObjectPaintChannels(
                    resources,
                    invocation.Name,
                    invocation.Transform,
                    invocation.ClipPath,
                    invocation.FillOpacity,
                    invocation.StrokeOpacity,
                    pageWidth,
                    pageHeight,
                    cache,
                    activeStreams,
                    pageContentBudget,
                    type3GlyphBudget,
                    depth + 1);
            }
            cache.VisibleForms[cacheKey] = channels;
            return channels;
        } catch (Exception exception) when (IsRecoverableType3ProjectionFailure(exception)) {
            return PdfType3PaintChannels.Both;
        } finally {
            activeStreams.Remove(form);
        }
    }

    private sealed class Type3PaintChannelCache {
        internal Dictionary<(
            PdfStream Stream,
            PdfDictionary Resources,
            Matrix2D ProgramTransform,
            PdfPageClipPath? ProgramClipPath,
            double? FillOpacity,
            double? StrokeOpacity,
            double PageWidth,
            double PageHeight), PdfType3PaintChannels> Streams { get; } =
            new Dictionary<(
                PdfStream Stream,
                PdfDictionary Resources,
                Matrix2D ProgramTransform,
                PdfPageClipPath? ProgramClipPath,
                double? FillOpacity,
                double? StrokeOpacity,
                double PageWidth,
                double PageHeight), PdfType3PaintChannels>();

        internal Dictionary<(
            PdfStream Stream,
            PdfDictionary Resources,
            Matrix2D InvocationTransform,
            PdfPageClipPath? InvocationClipPath,
            double? FillOpacity,
            double? StrokeOpacity,
            double PageWidth,
            double PageHeight), PdfType3PaintChannels> VisibleForms { get; } =
            new Dictionary<(
                PdfStream Stream,
                PdfDictionary Resources,
                Matrix2D InvocationTransform,
                PdfPageClipPath? InvocationClipPath,
                double? FillOpacity,
                double? StrokeOpacity,
                double PageWidth,
                double PageHeight), PdfType3PaintChannels>();
    }
}

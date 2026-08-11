using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private PdfType3PaintChannels ResolveType3PaintChannels(
        PdfPageType3GlyphInvocation glyph,
        Type3PaintChannelCache cache,
        HashSet<PdfStream> activeStreams,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        int depth = 0) {
        (double Width, double Height) visualPageSize = GetVisualPageSize();
        return ResolveType3PaintChannels(
            glyph,
            cache,
            activeStreams,
            pageContentBudget,
            type3GlyphBudget,
            visualPageSize.Width,
            visualPageSize.Height,
            depth);
    }

    private PdfType3PaintChannels ResolveType3PaintChannels(
        PdfPageType3GlyphInvocation glyph,
        Type3PaintChannelCache cache,
        HashSet<PdfStream> activeStreams,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        double pageWidth,
        double pageHeight,
        int depth) {
        if (glyph.Font.Type3 is not PdfType3FontResource type3 ||
            !type3.TryGetGlyph(glyph.CharacterCode, out PdfStream stream)) {
            return PdfType3PaintChannels.Both;
        }
        return ResolveType3PaintChannels(
            stream,
            type3.Resources,
            new PdfPageXObjectPaintState(
                Matrix2D.Multiply(glyph.Transform, type3.FontMatrix),
                glyph.ClipPath,
                glyph.FillOpacity,
                glyph.StrokeOpacity,
                glyph.StrokeWidth,
                glyph.StrokeDashStyle,
                glyph.StrokeLineCap,
                glyph.StrokeLineJoin),
            pageWidth,
            pageHeight,
            cache,
            activeStreams,
            pageContentBudget,
            type3GlyphBudget,
            depth);
    }

    private PdfType3PaintChannels ResolveVisibleFormPaintChannels(
        PdfStream form,
        PdfDictionary resources,
        PdfPageXObjectPaintState invocationState,
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
            InvocationState: invocationState,
            PageWidth: pageWidth,
            PageHeight: pageHeight);
        if (cache.VisibleForms.TryGetValue(cacheKey, out PdfType3PaintChannels cached)) return cached;
        if (!activeStreams.Add(form)) return PdfType3PaintChannels.Both;
        try {
            string content = WrapFormContentWithBoundingBoxClip(
                PdfEncoding.Latin1GetString(pageContentBudget.Decode(form)),
                form.Dictionary);
            Matrix2D formTransform = ApplyFormMatrix(invocationState.Transform, form.Dictionary);
            PdfType3PaintChannels channels = PdfType3PaintChannels.None;
            Dictionary<string, PdfPageColorSpace> colorSpaces = GetColorSpaceResources(resources);
            IReadOnlyDictionary<string, PdfPageGraphicsStateResource> graphicsStates = GetGraphicsStateResources(resources);
            IReadOnlyList<PdfPageDrawingEffectTransition> effects = PdfPageGraphicsEffectTimelineParser.Parse(
                content,
                graphicsStates,
                PdfPageDrawingEffect.Default,
                formTransform,
                maxOperations: _limits.MaxContentOperations,
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands);
            string transformedContent = WrapContentWithTransform(content, formTransform, out int transformedOffset);
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
                initialClipPath: invocationState.ClipPath,
                initialFillOpacity: invocationState.FillOpacity,
                initialStrokeOpacity: invocationState.StrokeOpacity,
                initialStrokeWidth: invocationState.StrokeWidth,
                initialStrokeDashStyle: invocationState.StrokeDashStyle,
                initialStrokeLineCap: invocationState.StrokeLineCap,
                initialStrokeLineJoin: invocationState.StrokeLineJoin,
                maxOperations: _limits.MaxContentOperations,
                patternBaseColorSpaces: GetPatternBaseColorSpaceResources(resources),
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                primitiveVisitor: primitive => {
                    PdfPageDrawingEffect effect = ResolveDrawingEffect(effects, primitive.PaintOrder);
                    if (IsPaintSuppressedByTransparentSoftMask(
                            effect,
                            resources,
                            formTransform,
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

            Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
            Dictionary<string, Func<byte[], double>> widthProviders = ResourceResolver.GetFontWidthProvidersForResources(resources, _objects);
            foreach (PdfPageXObjectInvocation invocation in PdfPageXObjectInvocationParser.Parse(
                         content,
                         formTransform,
                         pageHeight,
                         GetGraphicsStateResources(resources),
                         colorSpaces,
                         GetOptionalContentVisibility(resources),
                         initialClipPath: invocationState.ClipPath,
                         initialFillOpacity: invocationState.FillOpacity,
                         initialStrokeOpacity: invocationState.StrokeOpacity,
                         initialStrokeWidth: invocationState.StrokeWidth,
                         initialStrokeDashStyle: invocationState.StrokeDashStyle,
                         initialStrokeLineCap: invocationState.StrokeLineCap,
                         initialStrokeLineJoin: invocationState.StrokeLineJoin,
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
                         type3PaintChannelResolver: glyph => ResolveType3PaintChannels(
                             glyph,
                             cache,
                             activeStreams,
                             pageContentBudget,
                             type3GlyphBudget,
                             pageWidth,
                             pageHeight,
                             depth + 1),
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
                if (invocation.InlineImage != null &&
                    !IsInvisibleInlineImageInvocation(invocation, resources, pageWidth, pageHeight)) {
                    channels |= PdfType3PaintChannels.Fill;
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
            cache.VisibleForms[cacheKey] = channels;
            return channels;
        } catch (Exception exception) when (IsRecoverableType3ProjectionFailure(exception)) {
            return PdfType3PaintChannels.Both;
        } finally {
            activeStreams.Remove(form);
        }
    }

    private bool IsPaintSuppressedByTransparentSoftMask(
        PdfPageDrawingEffect effect,
        PdfDictionary? parentResources,
        Matrix2D fallbackTransform,
        double pageWidth,
        double pageHeight,
        Type3PaintChannelCache cache,
        HashSet<PdfStream> activeStreams,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        PdfPageSoftMaskResource? softMask = effect.SoftMask;
        return softMask != null && IsSoftMaskEntirelyTransparent(
            softMask,
            effect.SoftMaskTransform ?? fallbackTransform,
            parentResources,
            pageWidth,
            pageHeight,
            cache,
            activeStreams,
            pageContentBudget,
            type3GlyphBudget,
            depth);
    }

    private bool IsSoftMaskEntirelyTransparent(
        PdfPageSoftMaskResource softMask,
        Matrix2D transform,
        PdfDictionary? parentResources,
        double pageWidth,
        double pageHeight,
        Type3PaintChannelCache cache,
        HashSet<PdfStream> activeStreams,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        if (HasVisibleSoftMaskBackdrop(softMask)) return false;
        PdfDictionary maskResources = ResolveDictionary(
            softMask.Group.Dictionary.Items.TryGetValue("Resources", out PdfObject? value) ? value : null) ??
            parentResources ??
            new PdfDictionary();
        var maskState = new PdfPageXObjectPaintState(
            transform,
            clipPath: null,
            fillOpacity: null,
            strokeOpacity: null,
            strokeWidth: 1D,
            strokeDashStyle: OfficeStrokeDashStyle.Solid,
            strokeLineCap: null,
            strokeLineJoin: null);
        return ResolveVisibleFormPaintChannels(
            softMask.Group,
            maskResources,
            maskState,
            pageWidth,
            pageHeight,
            cache,
            activeStreams,
            pageContentBudget,
            type3GlyphBudget,
            depth) == PdfType3PaintChannels.None;
    }

    private static bool HasVisibleSoftMaskBackdrop(PdfPageSoftMaskResource softMask) =>
        softMask.Mode == OfficeSoftMaskMode.Alpha
            ? softMask.BackdropColor.A > 0
            : softMask.BackdropColor.A > 0 &&
              (softMask.BackdropColor.R > 0 ||
               softMask.BackdropColor.G > 0 ||
               softMask.BackdropColor.B > 0);

    private sealed class Type3PaintChannelCache {
        internal Dictionary<(
            PdfStream Stream,
            PdfDictionary Resources,
            PdfPageXObjectPaintState ProgramState,
            double PageWidth,
            double PageHeight), PdfType3PaintChannels> Streams { get; } =
            new Dictionary<(
                PdfStream Stream,
                PdfDictionary Resources,
                PdfPageXObjectPaintState ProgramState,
                double PageWidth,
                double PageHeight), PdfType3PaintChannels>();

        internal Dictionary<(
            PdfStream Stream,
            PdfDictionary Resources,
            PdfPageXObjectPaintState InvocationState,
            double PageWidth,
            double PageHeight), PdfType3PaintChannels> VisibleForms { get; } =
            new Dictionary<(
                PdfStream Stream,
                PdfDictionary Resources,
                PdfPageXObjectPaintState InvocationState,
                double PageWidth,
                double PageHeight), PdfType3PaintChannels>();
    }
}

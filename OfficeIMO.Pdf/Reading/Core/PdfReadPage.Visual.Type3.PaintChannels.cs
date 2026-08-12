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
                unsupportedShadingTransformVisitor: () => channels |= PdfType3PaintChannels.Both,
                requireExactType3ShadingProjection: true,
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
                                     depth);
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
                             depth),
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
                         visibleShadingVisitor: _ => channels |= PdfType3PaintChannels.Fill,
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
        // Transparency proof is only an optimization. A repeated group means
        // the mask graph is cyclic, so keep its paint visible and let strict
        // projection validation reject the cycle.
        if (!cache.ActiveSoftMaskTransparencyProofs.Add(softMask.Group)) return false;
        try {
            Type3SoftMaskValidationContext validation = type3GlyphBudget.GetOrCreateSoftMaskValidationContext(this);
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
            PdfType3PaintChannels channels = ResolveVisibleFormPaintChannels(
                softMask.Group,
                maskResources,
                maskState,
                pageWidth,
                pageHeight,
                cache,
                activeStreams,
                validation.TransparencyProofPageContentBudget,
                validation.TransparencyProofType3GlyphBudget,
                depth);
            if (channels == PdfType3PaintChannels.None) return true;
            return softMask.Mode == OfficeSoftMaskMode.Luminosity &&
                   IsLuminositySoftMaskEntirelyBlack(
                       softMask,
                       parentResources,
                       transform,
                       pageWidth,
                       pageHeight,
                       cache,
                       validation.TransparencyProofPageContentBudget,
                       validation.TransparencyProofType3GlyphBudget);
        } finally {
            cache.ActiveSoftMaskTransparencyProofs.Remove(softMask.Group);
        }
    }

    private bool IsLuminositySoftMaskEntirelyBlack(
        PdfPageSoftMaskResource softMask,
        PdfDictionary? parentResources,
        Matrix2D transform,
        double pageWidth,
        double pageHeight,
        Type3PaintChannelCache cache,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget) {
        PdfDictionary? effectiveParentResources = softMask.ParentResources ?? parentResources;
        var cacheKey = (softMask.Group, effectiveParentResources, transform, pageWidth, pageHeight);
        if (cache.BlackLuminosityForms.TryGetValue(cacheKey, out bool cached)) return cached;
        string content = PdfEncoding.Latin1GetString(pageContentBudget.Decode(softMask.Group));
        if (!IsVectorOnlyLuminosityProofContent(content)) {
            cache.BlackLuminosityForms[cacheKey] = false;
            return false;
        }
        var softMasks = new Dictionary<(PdfStream Group, PdfDictionary? ParentResources, OfficeSoftMaskMode Mode, OfficeColor Backdrop, Matrix2D Transform, double Width, double Height), OfficeDrawingSoftMask>();
        OfficeDrawing drawing = CreateFormDrawing(
            softMask.Group,
            effectiveParentResources,
            pageWidth,
            pageHeight,
            transform,
            softMasks,
            new HashSet<PdfStream>(),
            CreateTextOutputBudget(),
            pageContentBudget,
            type3GlyphBudget,
            decodedContent: content);
        bool result = IsEntirelyBlackLuminosityDrawing(drawing);
        cache.BlackLuminosityForms[cacheKey] = result;
        return result;
    }

    private bool IsVectorOnlyLuminosityProofContent(string content) {
        bool vectorOnly = true;
        PdfContentStreamInterpreter.InterpretUntil(
            content,
            _limits.MaxContentOperations,
            operation => {
                if (operation.Name is "BT" or "Do" or "BI") {
                    vectorOnly = false;
                    return false;
                }
                return true;
            },
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands);
        return vectorOnly;
    }

    private static bool IsEntirelyBlackLuminosityDrawing(OfficeDrawing drawing) {
        for (int index = 0; index < drawing.Elements.Count; index++) {
            OfficeDrawingElement element = drawing.Elements[index];
            switch (element) {
                case OfficeDrawingShape shape:
                    if (!IsEntirelyBlackLuminosityShape(shape.Shape)) return false;
                    break;
                case OfficeDrawingText:
                    // Ordinary text remains subject to strict font and text-mode validation.
                    // Do not let this optional black-luminance proof bypass those checks.
                    return false;
                case OfficeDrawingGroup group:
                    if (!IsEntirelyBlackLuminosityDrawing(group.Drawing)) return false;
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    if (effectGroup.Opacity > 0D && !IsEntirelyBlackLuminosityDrawing(effectGroup.Drawing)) return false;
                    break;
                default:
                    return false;
            }
        }
        return true;
    }

    private static bool IsEntirelyBlackLuminosityShape(OfficeShape shape) {
        if (shape.Shadow != null || shape.Glow != null ||
            shape.FillGradient != null || shape.FillRadialGradient != null ||
            shape.StrokeGradient != null || shape.StrokeRadialGradient != null) {
            return false;
        }
        if ((shape.FillOpacity ?? 1D) > 0D && !IsBlackOrTransparent(shape.FillColor)) return false;
        return shape.StrokeWidth <= 0D ||
               (shape.StrokeOpacity ?? 1D) <= 0D ||
               IsBlackOrTransparent(shape.StrokeColor);
    }

    private static bool IsBlackOrTransparent(OfficeColor? color) =>
        !color.HasValue || color.Value.A == 0 ||
        (color.Value.R == 0 && color.Value.G == 0 && color.Value.B == 0);

    private static bool HasVisibleSoftMaskBackdrop(PdfPageSoftMaskResource softMask) =>
        softMask.Mode == OfficeSoftMaskMode.Alpha
            ? softMask.BackdropColor.A > 0
            : softMask.BackdropColor.A > 0 &&
              (softMask.BackdropColor.R > 0 ||
               softMask.BackdropColor.G > 0 ||
               softMask.BackdropColor.B > 0);

    private sealed class Type3PaintChannelCache {
        internal HashSet<PdfStream> ActiveSoftMaskTransparencyProofs { get; } = new HashSet<PdfStream>();

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
            PdfDictionary? ParentResources,
            Matrix2D Transform,
            double PageWidth,
            double PageHeight), bool> BlackLuminosityForms { get; } =
            new Dictionary<(
                PdfStream Stream,
                PdfDictionary? ParentResources,
                Matrix2D Transform,
                double PageWidth,
                double PageHeight), bool>();

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

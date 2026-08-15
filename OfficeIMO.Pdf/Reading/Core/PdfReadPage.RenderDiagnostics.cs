using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal IReadOnlyList<PdfRenderCapabilityDiagnostic> GetRenderCapabilityDiagnostics() {
        var diagnostics = new List<PdfRenderCapabilityDiagnostic>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        PdfOutputIntentColorTransform? outputIntentColorTransform = _outputIntentColorTransform;
        if (outputIntentColorTransform != null) {
            if (!outputIntentColorTransform.IsSupported) {
                AddRenderDiagnostic(
                    diagnostics,
                    seen,
                    PdfRenderCapabilities.UnsupportedIccOutputIntentId,
                    outputIntentColorTransform.Subject);
            } else if (_hasOutputIntentCompositionInteraction?.Value == true) {
                AddRenderDiagnostic(
                    diagnostics,
                    seen,
                    PdfRenderCapabilities.OutputIntentTransparencyId,
                    outputIntentColorTransform.Subject);
            }
        }
        var activeForms = new HashSet<PdfStream>();
        var pageContentBudget = new PageContentBudget(this);
        var type3GlyphBudget = new Type3GlyphBudget(_limits.MaxType3GlyphInvocationsPerPage);
        var textClippingBudget = new PdfTextClippingBudget();
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        CollectRenderCapabilityDiagnostics(
            GetContentStreamContent(pageContentBudget),
            resources,
            diagnostics,
            seen,
            activeForms,
            pageContentBudget,
            type3GlyphBudget,
            textClippingBudget,
            0,
            0,
            GetVisualPageTransform());
        CollectAnnotationCapabilityDiagnostics(diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, textClippingBudget);
        return diagnostics.Count == 0 ? Array.Empty<PdfRenderCapabilityDiagnostic>() : diagnostics.AsReadOnly();
    }

    private void CollectRenderCapabilityDiagnostics(
        string content,
        PdfDictionary? resources,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        PdfTextClippingBudget textClippingBudget,
        int depth,
        int auxiliarySurfaceDepth,
        Matrix2D? initialTransform = null,
        PdfPageColorSpace initialFillColorSpace = default,
        PdfPagePatternSelection? initialFillPattern = null,
        PdfPageColorSpace? initialFillPatternBaseColorSpace = null,
        PdfPageColorSpace initialStrokeColorSpace = default,
        PdfPagePatternSelection? initialStrokePattern = null,
        PdfPageColorSpace? initialStrokePatternBaseColorSpace = null,
        PdfPageClipPath? initialClipPath = null) {
        EnsureContentNestingBudget(depth);
        PdfPageInvokedResourceNames invokedResources = GetInvokedResourceNames(content, resources);
        HashSet<string> unsupportedColorSpaces = GetUnsupportedColorSpaceResourceNames(resources, pageContentBudget, invokedResources.ColorSpaces);
        HashSet<string> approximatedIccColorSpaces = GetApproximatedIccColorSpaceResourceNames(resources, pageContentBudget, invokedResources.ColorSpaces);
        var invokedXObjects = new List<string>();
        var invokedFonts = new HashSet<string>(StringComparer.Ordinal);
        var invokedShadings = new HashSet<string>(StringComparer.Ordinal);
        var invokedPatterns = new HashSet<string>(StringComparer.Ordinal);
        var invokedSoftMasks = new HashSet<PdfStream>();
        var invokedXObjectStates = new List<PdfPageXObjectInvocation>();
        PdfContentStreamInterpreter.Interpret(content, _limits.MaxContentOperations, operation => {
            string? capabilityId = GetOperatorCapabilityId(operation.Name);
            if (capabilityId != null) AddRenderDiagnostic(diagnostics, seen, capabilityId, operation.Name);
            if (operation.Name == "Do" &&
                operation.Operands.Count > 0 &&
                operation.Operands[operation.Operands.Count - 1] is string xObjectName) {
                invokedXObjects.Add(xObjectName);
            }
            if (operation.Name == "sh" && operation.Operands.Count > 0 &&
                operation.Operands[operation.Operands.Count - 1] is string shadingName) {
                invokedShadings.Add(shadingName);
            }
            if ((operation.Name == "scn" || operation.Name == "SCN") &&
                operation.Operands.Count > 0 &&
                operation.Operands[operation.Operands.Count - 1] is string patternName) {
                invokedPatterns.Add(patternName);
            }
            if ((operation.Name == "cs" || operation.Name == "CS") &&
                operation.Operands.Count > 0 &&
                operation.Operands[operation.Operands.Count - 1] is string colorSpaceName) {
                if (unsupportedColorSpaces.Contains(colorSpaceName)) {
                    AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.ColorSpaceId, colorSpaceName);
                } else if (approximatedIccColorSpaces.Contains(colorSpaceName)) {
                    AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.IccColorSpaceId, colorSpaceName);
                }
            }
            if (operation.InlineImage is PdfContentInlineImage inlineImage) {
                CollectImageColorSpaceCapabilityDiagnostic(
                    inlineImage.Dictionary,
                    resources,
                    diagnostics,
                    seen,
                    "inline-image",
                    pageContentBudget,
                    new PdfStream(inlineImage.Dictionary, inlineImage.Data));
            }
        },
        maxNestingDepth: _limits.MaxContentNestingDepth,
        maxOperands: _limits.MaxContentOperands,
        inlineImageArrayComponentCount: array => GetDeclaredColorSpaceComponentCount(array));

        if (resources == null) return;
        if (invokedShadings.Count > 0 || invokedPatterns.Count > 0) {
            _ = PdfPageContentVisualParser.Parse(
                WrapContentWithTransform(content, initialTransform ?? Matrix2D.Identity),
                GetVisualPageSize().Width,
                GetVisualPageSize().Height,
                GetGraphicsStateResources(resources),
                GetColorSpaceResources(resources, pageContentBudget: pageContentBudget),
                GetShadingResources(resources, pageContentBudget: pageContentBudget),
                GetShadingPatternResources(resources, pageContentBudget: pageContentBudget),
                tilingPatterns: null,
                GetOptionalContentVisibility(resources),
                initialFillColorSpace: initialFillColorSpace,
                initialStrokeColorSpace: initialStrokeColorSpace,
                maxOperations: _limits.MaxContentOperations,
                patternBaseColorSpaces: GetPatternBaseColorSpaceResources(resources, pageContentBudget: pageContentBudget),
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                primitiveVisitor: static _ => { },
                retainPrimitiveData: false,
                unsupportedShadingTransformVisitor: () => AddRenderDiagnostic(
                    diagnostics,
                    seen,
                    PdfRenderCapabilities.UnsupportedShadingId,
                    "transformed-radial-shading"),
                textClippingBudget: textClippingBudget,
                inlineImageArrayComponentCount: array => GetDeclaredColorSpaceComponentCount(array));
        }
        HashSet<string> failedType3Fonts = CollectType3FontFailures(
            content,
            resources,
            initialTransform ?? Matrix2D.Identity,
            initialFillColorSpace,
            initialFillPattern,
            initialFillPatternBaseColorSpace,
            initialStrokeColorSpace,
            initialStrokePattern,
            initialStrokePatternBaseColorSpace,
            initialClipPath,
            pageContentBudget,
            type3GlyphBudget,
            textClippingBudget,
            invokedFonts,
            invokedPatterns,
            invokedSoftMasks,
            invokedXObjectStates,
            diagnostics,
            seen,
            depth);
        CollectFontCapabilityDiagnostics(resources, invokedFonts, failedType3Fonts, diagnostics, seen);
        CollectShadingCapabilityDiagnostics(resources, invokedShadings, invokedPatterns, diagnostics, seen, pageContentBudget);
        CollectPatternCapabilityDiagnostics(resources, diagnostics, seen);
        CollectGraphicsStateCapabilityDiagnostics(resources, diagnostics, seen);
        CollectXObjectCapabilityDiagnostics(resources, invokedXObjects, invokedXObjectStates, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, textClippingBudget, depth, auxiliarySurfaceDepth);
        CollectAuxiliarySurfaceCapabilityDiagnostics(resources, invokedPatterns, invokedSoftMasks, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, textClippingBudget, auxiliarySurfaceDepth);
    }

    private static string? GetOperatorCapabilityId(string op) {
        switch (op) {
            case "M": return PdfRenderCapabilities.MiterLimitId;
            case "ri": return null;
            case "i": return PdfRenderCapabilities.FlatnessId;
            case "MP":
            case "DP": return PdfRenderCapabilities.MarkedPointId;
            case "d0":
            case "d1": return PdfRenderCapabilities.Type3MetricsId;
            default: return PdfContentOperators.IsStandard(op) ? null : PdfRenderCapabilities.UnknownOperatorId;
        }
    }

    private HashSet<string> CollectType3FontFailures(
        string content,
        PdfDictionary resources,
        Matrix2D initialTransform,
        PdfPageColorSpace initialFillColorSpace,
        PdfPagePatternSelection? initialFillPattern,
        PdfPageColorSpace? initialFillPatternBaseColorSpace,
        PdfPageColorSpace initialStrokeColorSpace,
        PdfPagePatternSelection? initialStrokePattern,
        PdfPageColorSpace? initialStrokePatternBaseColorSpace,
        PdfPageClipPath? initialClipPath,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        PdfTextClippingBudget textClippingBudget,
        HashSet<string> invokedFonts,
        HashSet<string> invokedPatterns,
        HashSet<PdfStream> invokedSoftMasks,
        List<PdfPageXObjectInvocation> invokedXObjectStates,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        int contentNestingDepth) {
        var failures = new HashSet<string>(StringComparer.Ordinal);
        var activeStreams = new HashSet<PdfStream>();
        PdfPageInvokedResourceNames invokedResources = GetInvokedResourceNames(content, resources);
        Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
        Dictionary<string, Func<byte[], double>> widthProviders = ResourceResolver.GetFontWidthProvidersForResources(resources, _objects);
        Dictionary<string, PdfPageColorSpace> colorSpaces = GetColorSpaceResources(resources, invokedResources.ColorSpaces, pageContentBudget);
        Dictionary<string, PdfPageColorSpace> patternBaseColorSpaces = GetPatternBaseColorSpaceResources(resources, invokedResources.ColorSpaces, pageContentBudget);
        var invokedPatternNames = new HashSet<string>(StringComparer.Ordinal);
        IReadOnlyList<PdfPageXObjectInvocation> discoveredXObjects = PdfPageXObjectInvocationParser.Parse(
            content,
            initialTransform,
            GetVisualPageSize().Height,
            GetGraphicsStateResources(resources),
            colorSpaces,
            GetOptionalContentVisibility(resources),
            maxOperations: _limits.MaxContentOperations,
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            initialClipPath: initialClipPath,
            initialFillColorSpace: initialFillColorSpace,
            initialStrokeColorSpace: initialStrokeColorSpace,
            fonts: fonts,
            fontWidthProviders: widthProviders,
            visibleFontVisitor: fontName => {
                if (!string.IsNullOrEmpty(fontName)) invokedFonts.Add(fontName);
            },
            patternInvocationVisitor: name => invokedPatternNames.Add(name),
            graphicsStateVisitor: (state, _, _, _, _, _) => {
                if (state.SoftMask?.Group is PdfStream group) invokedSoftMasks.Add(group);
            },
            patternBaseColorSpaces: patternBaseColorSpaces,
            initialFillPattern: initialFillPattern,
            initialFillPatternBaseColorSpace: initialFillPatternBaseColorSpace,
            initialStrokePattern: initialStrokePattern,
            initialStrokePatternBaseColorSpace: initialStrokePatternBaseColorSpace,
            textClippingBudget: textClippingBudget);
        bool invokesType3 = invokedFonts.Any(name => fonts.TryGetValue(name, out PdfFontResource? font) && font.Type3 != null);
        bool carriesPatternIntoXObject = discoveredXObjects.Any(invocation => invocation.FillPattern.HasValue || invocation.StrokePattern.HasValue);
        if (!invokesType3 && !carriesPatternIntoXObject) {
            invokedXObjectStates.AddRange(discoveredXObjects);
            invokedPatterns.UnionWith(invokedPatternNames);
            return failures;
        }
        Dictionary<string, PdfPageTilingPatternResource> tilingPatterns = GetTilingPatternResources(
            resources,
            invokedPatternNames,
            textOutputBudget: type3GlyphBudget.GetOrCreateSoftMaskValidationContext(
                this,
                pageContentBudget,
                textClippingBudget,
                textClippingBudget).TextOutputBudget,
            pageContentBudget: pageContentBudget,
            type3GlyphBudget: type3GlyphBudget,
            requireSupportedType3Content: false,
            contentNestingDepth: contentNestingDepth,
            invocationTextClippingBudget: textClippingBudget,
            patternTextClippingBudget: textClippingBudget);
        IReadOnlyList<PdfPageXObjectInvocation> resolvedXObjects = PdfPageXObjectInvocationParser.Parse(
            content,
            initialTransform,
            GetVisualPageSize().Height,
            GetGraphicsStateResources(resources),
            colorSpaces,
            GetOptionalContentVisibility(resources),
            maxOperations: _limits.MaxContentOperations,
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            initialClipPath: initialClipPath,
            initialFillColorSpace: initialFillColorSpace,
            initialStrokeColorSpace: initialStrokeColorSpace,
            fonts: fonts,
            fontWidthProviders: widthProviders,
            type3TextVisitor: invocation => {
                bool supported = true;
                for (int i = 0; i < invocation.Glyphs.Count; i++) {
                    PdfPageType3GlyphInvocation glyph = invocation.Glyphs[i];
                    if (glyph.Font.Type3 is not PdfType3FontResource type3 ||
                        !type3.TryGetGlyph(glyph.CharacterCode, out PdfStream stream) ||
                        !CanProjectType3GlyphProgram(
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
                            type3.IsUncolored,
                            glyph.FillPattern,
                            glyph.FillPatternBaseColorSpace,
                            glyph.StrokePattern,
                            glyph.StrokePatternBaseColorSpace,
                            pageContentBudget,
                            type3GlyphBudget,
                            textClippingBudget,
                            activeStreams,
                            diagnostics,
                            seen,
                            contentNestingDepth + 1,
                            initialHasAuthoredRenderingIntent: glyph.HasAuthoredRenderingIntent,
                            initialRenderingIntent: glyph.RenderingIntent,
                            initialFillColorSelection: glyph.FillColorSelection,
                            initialStrokeColorSelection: glyph.StrokeColorSelection)) {
                        failures.Add(glyph.Font.ResourceName);
                        supported = false;
                    }
                }
                return supported;
            },
            type3GlyphBudgetConsumer: type3GlyphBudget.Consume,
            visibleFontVisitor: fontName => {
                if (!string.IsNullOrEmpty(fontName)) invokedFonts.Add(fontName);
            },
            patternInvocationVisitor: patternName => invokedPatterns.Add(patternName),
            graphicsStateVisitor: (state, _, _, _, _, _) => {
                if (state.SoftMask?.Group is PdfStream group) invokedSoftMasks.Add(group);
            },
            patternBaseColorSpaces: patternBaseColorSpaces,
            initialFillPattern: initialFillPattern,
            initialFillPatternBaseColorSpace: initialFillPatternBaseColorSpace,
            initialStrokePattern: initialStrokePattern,
            initialStrokePatternBaseColorSpace: initialStrokePatternBaseColorSpace,
            tilingPatterns: tilingPatterns,
            shadingPatterns: GetShadingPatternResources(resources, pageContentBudget: pageContentBudget),
            textClippingBudget: textClippingBudget);
        invokedXObjectStates.AddRange(resolvedXObjects);
        return failures;
    }

    private bool CanProjectType3GlyphProgram(
        PdfStream stream,
        PdfDictionary resources,
        PdfPageXObjectPaintState programState,
        bool requireImageMask,
        PdfPagePatternSelection? initialFillPattern,
        PdfPageColorSpace? initialFillPatternBaseColorSpace,
        PdfPagePatternSelection? initialStrokePattern,
        PdfPageColorSpace? initialStrokePatternBaseColorSpace,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        PdfTextClippingBudget textClippingBudget,
        HashSet<PdfStream> activeStreams,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        int depth,
        SoftMaskNestingDepth? softMaskNestingDepth = null,
        Dictionary<(PdfStream Group, PdfDictionary? ParentResources, Matrix2D Transform, double Width, double Height), int>? validatedSoftMaskGroups = null,
        HashSet<PdfStream>? activeSoftMaskGroups = null,
        HashSet<PdfStream>? activeSoftMaskForms = null,
        double? projectionPageWidth = null,
        double? projectionPageHeight = null,
        bool requireIsolatedGroupSemantics = false,
        bool initialHasAuthoredRenderingIntent = false,
        OfficeIccRenderingIntent initialRenderingIntent = OfficeIccRenderingIntent.RelativeColorimetric,
        PdfPaintColorSelection? initialFillColorSelection = null,
        PdfPaintColorSelection? initialStrokeColorSelection = null) {
        EnsureContentNestingBudget(depth);
        if (softMaskNestingDepth != null) {
            softMaskNestingDepth.Maximum = Math.Max(softMaskNestingDepth.Maximum, depth);
        }
        if (Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, _objects).Count != 0 ||
            !activeStreams.Add(stream)) return false;
        try {
            string content;
            try {
                content = PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream));
            } catch (IOException exception) when (exception is not PdfReadLimitException) {
                return false;
            }

            (double Width, double Height) visualPageSize = GetVisualPageSize();
            double surfaceWidth = projectionPageWidth ?? visualPageSize.Width;
            double surfaceHeight = projectionPageHeight ?? visualPageSize.Height;
            bool supported = true;
            if (initialFillPattern.HasValue && initialFillPatternBaseColorSpace?.UsesIccApproximation == true) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.IccColorSpaceId, initialFillPattern.Value.Name);
            }
            if (initialStrokePattern.HasValue && initialStrokePatternBaseColorSpace?.UsesIccApproximation == true) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.IccColorSpaceId, initialStrokePattern.Value.Name);
            }
            validatedSoftMaskGroups ??= new Dictionary<(PdfStream Group, PdfDictionary? ParentResources, Matrix2D Transform, double Width, double Height), int>();
            activeSoftMaskGroups ??= new HashSet<PdfStream>();
            activeSoftMaskForms ??= new HashSet<PdfStream>();
            softMaskNestingDepth ??= new SoftMaskNestingDepth(depth);
            Type3SoftMaskValidationContext softMaskValidation =
                type3GlyphBudget.GetOrCreateSoftMaskValidationContext(
                    this,
                    pageContentBudget,
                    textClippingBudget,
                    textClippingBudget);
            Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
            Dictionary<string, Func<byte[], double>> widthProviders = ResourceResolver.GetFontWidthProvidersForResources(resources, _objects);
            Dictionary<string, PdfPageColorSpace> colorSpaces = GetColorSpaceResources(resources, pageContentBudget: pageContentBudget);
            Dictionary<string, PdfPageColorSpace> patternBaseColorSpaces = GetPatternBaseColorSpaceResources(resources, pageContentBudget: pageContentBudget);
            IReadOnlyDictionary<string, PdfPageGraphicsStateResource> graphicsStates = GetGraphicsStateResources(resources);
            IReadOnlyList<PdfPageDrawingEffectTransition> drawingEffects = PdfPageGraphicsEffectTimelineParser.Parse(
                content,
                graphicsStates,
                PdfPageDrawingEffect.Default,
                programState.Transform,
                maxOperations: _limits.MaxContentOperations,
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                inlineImageArrayComponentCount: array => GetDeclaredColorSpaceComponentCount(array));
            if (!HasSupportedType3PageBlendColorSpace() &&
                drawingEffects.Any(static transition => transition.Effect.BlendMode != OfficeBlendMode.Normal)) return false;
            if (requireIsolatedGroupSemantics &&
                drawingEffects.Any(static transition => transition.Effect.BlendMode != OfficeBlendMode.Normal)) return false;
            var invokedPatternNames = new HashSet<string>(StringComparer.Ordinal);
            var type3PaintChannelCache = new Type3PaintChannelCache();
            var activeType3PaintChannelStreams = new HashSet<PdfStream>();
            _ = PdfPageXObjectInvocationParser.Parse(
                content,
                programState.Transform,
                surfaceHeight,
                graphicsStates,
                colorSpaces,
                GetOptionalContentVisibility(resources),
                maxOperations: _limits.MaxContentOperations,
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                initialClipPath: programState.ClipPath,
                initialFillOpacity: programState.FillOpacity,
                initialStrokeOpacity: programState.StrokeOpacity,
                initialStrokeWidth: programState.StrokeWidth,
                initialStrokeDashStyle: programState.StrokeDashStyle,
                initialStrokeLineCap: programState.StrokeLineCap,
                initialStrokeLineJoin: programState.StrokeLineJoin,
                fonts: fonts,
                fontWidthProviders: widthProviders,
                patternInvocationVisitor: name => invokedPatternNames.Add(name),
                patternBaseColorSpaces: patternBaseColorSpaces,
                initialFillPattern: initialFillPattern,
                initialFillPatternBaseColorSpace: initialFillPatternBaseColorSpace,
                initialStrokePattern: initialStrokePattern,
                initialStrokePatternBaseColorSpace: initialStrokePatternBaseColorSpace,
                type3PaintChannelResolver: glyph => ResolveType3PaintChannels(
                    glyph,
                    type3PaintChannelCache,
                    activeType3PaintChannelStreams,
                    pageContentBudget,
                    type3GlyphBudget),
                xObjectPaintChannelResolver: (name, paintState) => ResolveXObjectPaintChannels(
                    resources,
                    name,
                    paintState,
                    surfaceWidth,
                    surfaceHeight,
                    type3PaintChannelCache,
                    activeType3PaintChannelStreams,
                    pageContentBudget,
                    type3GlyphBudget),
                softMaskVisibilityResolver: (softMask, transform, fillColor, strokeColor, hasFillPattern, hasStrokePattern) =>
                    LuminositySoftMaskDependsOnInheritedPaint(softMask, fillColor, strokeColor, hasFillPattern, hasStrokePattern, pageContentBudget) ||
                    !IsSoftMaskEntirelyTransparent(
                        softMask,
                        transform,
                        resources,
                        surfaceWidth,
                        surfaceHeight,
                        type3PaintChannelCache,
                        activeType3PaintChannelStreams,
                        pageContentBudget,
                        type3GlyphBudget,
                        depth + 1),
                pageWidth: surfaceWidth,
                textClippingBudget: textClippingBudget,
                initialHasAuthoredRenderingIntent: initialHasAuthoredRenderingIntent,
                initialRenderingIntent: initialRenderingIntent,
                initialFillColorSelection: initialFillColorSelection,
                initialStrokeColorSelection: initialStrokeColorSelection,
                outputIntentColorTransform: EffectiveOutputIntentColorTransform);
            if (invokedPatternNames.Count > 0 && softMaskNestingDepth != null) {
                softMaskNestingDepth.Cacheable = false;
            }
            Dictionary<string, PdfPageTilingPatternResource> tilingPatterns = GetTilingPatternResources(
                resources,
                invokedPatternNames,
                textOutputBudget: softMaskValidation.TextOutputBudget,
                pageContentBudget: pageContentBudget,
                type3GlyphBudget: type3GlyphBudget,
                requireSupportedType3Content: true,
                contentNestingDepth: depth,
                invocationTextClippingBudget: textClippingBudget,
                patternTextClippingBudget: textClippingBudget);
            Dictionary<string, PdfPageShadingPatternResource> shadingPatterns = GetShadingPatternResources(resources, pageContentBudget: pageContentBudget);
            Dictionary<string, PdfPageShadingResource> directShadings = GetShadingResources(resources, pageContentBudget: pageContentBudget);
            bool usesFillPaint = false;
            bool usesStrokePaint = false;
            bool usesUnsupportedInheritedShadingStroke = false;
            bool usesUnsupportedInheritedShadingPlacement = false;
            _ = PdfPageContentVisualParser.Parse(
                WrapContentWithTransform(content, programState.Transform),
                surfaceWidth,
                surfaceHeight,
                graphicsStates,
                colorSpaces,
                directShadings,
                shadingPatterns,
                tilingPatterns,
                GetOptionalContentVisibility(resources),
                maxOperations: _limits.MaxContentOperations,
                patternBaseColorSpaces: patternBaseColorSpaces,
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                initialClipPath: programState.ClipPath,
                initialFillOpacity: programState.FillOpacity,
                initialStrokeOpacity: programState.StrokeOpacity,
                initialStrokeWidth: programState.StrokeWidth,
                initialStrokeDashStyle: programState.StrokeDashStyle,
                initialStrokeLineCap: programState.StrokeLineCap,
                initialStrokeLineJoin: programState.StrokeLineJoin,
                scaleStrokeWidthWithTransform: true,
                primitiveVisitor: primitive => {
                    if (!CanRenderTilingPatterns(primitive, surfaceWidth, surfaceHeight)) supported = false;
                    usesFillPaint |= primitive.HasFillPaint;
                    usesStrokePaint |= primitive.HasStrokePaint;
                    usesUnsupportedInheritedShadingStroke |=
                        primitive.HasStrokePaint &&
                        initialStrokePattern?.ShadingPattern.HasValue == true;
                    if (primitive.HasFillPaint && initialFillPattern?.ShadingPattern.HasValue == true) {
                        PdfPageShadingPatternResource pattern = initialFillPattern.Value.ShadingPattern.Value;
                        Matrix2D combined = Matrix2D.Multiply(initialFillPattern.Value.PaintTransform, pattern.Matrix);
                        usesUnsupportedInheritedShadingPlacement |=
                            !PdfPageContentVisualParser.IsSupportedExactShadingPlacement(
                                pattern.Shading,
                                combined,
                                primitive.X,
                                primitive.Y,
                                primitive.Width,
                                primitive.Height,
                                surfaceHeight);
                    }
                },
                retainPrimitiveData: false,
                unsupportedShadingTransformVisitor: () => {
                    supported = false;
                    AddRenderDiagnostic(
                        diagnostics,
                        seen,
                        PdfRenderCapabilities.UnsupportedShadingId,
                        "type3-shading-paint");
                },
                requireExactType3ShadingProjection: true,
                authoredShadingInvocationVisitor: requireImageMask
                    ? _ => supported = false
                    : null,
                unsupportedOperatorVisitor: _ => supported = false,
                initialFillPattern: initialFillPattern,
                initialStrokePattern: initialStrokePattern,
                textClippingBudget: textClippingBudget,
                inlineImageArrayComponentCount: array => GetDeclaredColorSpaceComponentCount(array));
            if (usesUnsupportedInheritedShadingStroke) {
                AddRenderDiagnostic(
                    diagnostics,
                    seen,
                    PdfRenderCapabilities.UnsupportedShadingId,
                    initialStrokePattern?.Name ?? "type3-shading-stroke");
                return false;
            }
            if (usesUnsupportedInheritedShadingPlacement) {
                AddRenderDiagnostic(
                    diagnostics,
                    seen,
                    PdfRenderCapabilities.UnsupportedShadingId,
                    initialFillPattern?.Name ?? "type3-shading-placement");
                return false;
            }
            if ((usesFillPaint && initialFillPattern.HasValue && !HasUsableInheritedPattern(initialFillPattern)) ||
                (usesStrokePaint && initialStrokePattern.HasValue && !HasUsableInheritedPattern(initialStrokePattern))) {
                PdfPagePatternSelection? unsupported = usesFillPaint && initialFillPattern.HasValue && !HasUsableInheritedPattern(initialFillPattern)
                    ? initialFillPattern
                    : initialStrokePattern;
                if (unsupported.HasValue && unsupported.Value.ShadingPattern.HasValue) {
                    AddRenderDiagnostic(
                        diagnostics,
                        seen,
                        PdfRenderCapabilities.UnsupportedShadingId,
                        unsupported.Value.Name);
                }
                return false;
            }
            var patternSupport = new Dictionary<string, bool>(StringComparer.Ordinal);
            foreach (PdfPageXObjectInvocation invocation in PdfPageXObjectInvocationParser.Parse(
                         content,
                         programState.Transform,
                         surfaceHeight,
                         graphicsStates,
                         colorSpaces,
                         GetOptionalContentVisibility(resources),
                         maxOperations: _limits.MaxContentOperations,
                         maxNestingDepth: _limits.MaxContentNestingDepth,
                         maxOperands: _limits.MaxContentOperands,
                         initialClipPath: programState.ClipPath,
                         initialFillOpacity: programState.FillOpacity,
                         initialStrokeOpacity: programState.StrokeOpacity,
                         initialStrokeWidth: programState.StrokeWidth,
                         initialStrokeDashStyle: programState.StrokeDashStyle,
                         initialStrokeLineCap: programState.StrokeLineCap,
                         initialStrokeLineJoin: programState.StrokeLineJoin,
                         fonts: fonts,
                         fontWidthProviders: widthProviders,
                         type3TextVisitor: nested => {
                             for (int index = 0; index < nested.Glyphs.Count; index++) {
                                 PdfPageType3GlyphInvocation glyph = nested.Glyphs[index];
                                 if (glyph.Font.Type3 is not PdfType3FontResource nestedType3 ||
                                     requireImageMask && !nestedType3.IsUncolored ||
                                     !nestedType3.TryGetGlyph(glyph.CharacterCode, out PdfStream nestedStream) ||
                                     (requireImageMask && !nestedType3.IsUncolored) ||
                                     !CanProjectType3GlyphProgram(
                                         nestedStream,
                                         nestedType3.Resources,
                                         new PdfPageXObjectPaintState(
                                             Matrix2D.Multiply(glyph.Transform, nestedType3.FontMatrix),
                                             glyph.ClipPath,
                                             glyph.FillOpacity,
                                             glyph.StrokeOpacity,
                                             glyph.StrokeWidth,
                                             glyph.StrokeDashStyle,
                                             glyph.StrokeLineCap,
                                             glyph.StrokeLineJoin),
                                         requireImageMask || nestedType3.IsUncolored,
                                         glyph.FillPattern,
                                         glyph.FillPatternBaseColorSpace,
                                         glyph.StrokePattern,
                                         glyph.StrokePatternBaseColorSpace,
                                         pageContentBudget,
                                         type3GlyphBudget,
                                         textClippingBudget,
                                         activeStreams,
                                         diagnostics,
                                         seen,
                                         depth + 1,
                                         softMaskNestingDepth,
                                         validatedSoftMaskGroups,
                                         activeSoftMaskGroups,
                                         activeSoftMaskForms,
                                         surfaceWidth,
                                         surfaceHeight,
                                         initialHasAuthoredRenderingIntent: glyph.HasAuthoredRenderingIntent,
                                         initialRenderingIntent: glyph.RenderingIntent,
                                         initialFillColorSelection: glyph.FillColorSelection,
                                         initialStrokeColorSelection: glyph.StrokeColorSelection)) {
                                     supported = false;
                                 }
                             }
                             return supported;
                         },
                         type3GlyphBudgetConsumer: type3GlyphBudget.Consume,
                         unsupportedTextVisitor: () => supported = false,
                         unsupportedGraphicsEffectVisitor: () => supported = false,
                         allowSupportedGraphicsEffects: true,
                         graphicsStateVisitor: (resource, resourceTransform, fillColor, strokeColor, hasFillPattern, hasStrokePattern) => {
                             if (!CanDecodeType3SoftMask(
                                     resource.SoftMask,
                                     resourceTransform,
                                     pageContentBudget,
                                     validatedSoftMaskGroups,
                                     type3GlyphBudget,
                                     activeSoftMaskGroups,
                                     activeSoftMaskForms,
                                     activeStreams,
                                     depth + 1,
                                     softMaskNestingDepth!,
                                     surfaceWidth,
                                     surfaceHeight,
                                     softMaskValidation.TextOutputBudget,
                                     fillColor,
                                     strokeColor,
                                     hasFillPattern,
                                     hasStrokePattern,
                                     resource)) {
                                 supported = false;
                             }
                         },
                         unsupportedColorVisitor: () => supported = false,
                         unsupportedPatternVisitor: requireImageMask ? () => supported = false : null,
                         patternInvocationVisitor: name => {
                             if (!patternSupport.TryGetValue(name, out bool canProject)) {
                                 canProject = false;
                                 if (requireImageMask &&
                                     ((initialFillPattern.HasValue && string.Equals(initialFillPattern.Value.Name, name, StringComparison.Ordinal)) ||
                                      (initialStrokePattern.HasValue && string.Equals(initialStrokePattern.Value.Name, name, StringComparison.Ordinal)))) {
                                     canProject = true;
                                } else if (!requireImageMask &&
                                           shadingPatterns.TryGetValue(name, out PdfPageShadingPatternResource shadingPattern)) {
                                    canProject = shadingPattern.SupportsExactType3Projection;
                                     CollectShadingCapabilityDiagnostics(
                                         resources,
                                         Array.Empty<string>(),
                                         new[] { name! },
                                         diagnostics,
                                         seen,
                                         pageContentBudget);
                                 } else if (!requireImageMask) {
                                     canProject = tilingPatterns.ContainsKey(name);
                                 }
                                 patternSupport[name] = canProject;
                             }
                             if (!canProject) supported = false;
                         },
                         patternBaseColorSpaces: patternBaseColorSpaces,
                         initialFillPattern: initialFillPattern,
                         initialFillPatternBaseColorSpace: initialFillPatternBaseColorSpace,
                         initialStrokePattern: initialStrokePattern,
                         initialStrokePatternBaseColorSpace: initialStrokePatternBaseColorSpace,
                         tilingPatterns: tilingPatterns,
                         shadingPatterns: shadingPatterns,
                         type3PaintChannelResolver: glyph => ResolveType3PaintChannels(
                             glyph,
                             type3PaintChannelCache,
                             activeType3PaintChannelStreams,
                             pageContentBudget,
                             type3GlyphBudget,
                             depth),
                         xObjectPaintChannelResolver: (name, paintState) => ResolveXObjectPaintChannels(
                             resources,
                             name,
                             paintState,
                             surfaceWidth,
                             surfaceHeight,
                             type3PaintChannelCache,
                             activeType3PaintChannelStreams,
                             pageContentBudget,
                             type3GlyphBudget),
                         softMaskVisibilityResolver: (softMask, transform, fillColor, strokeColor, hasFillPattern, hasStrokePattern) =>
                             LuminositySoftMaskDependsOnInheritedPaint(softMask, fillColor, strokeColor, hasFillPattern, hasStrokePattern, pageContentBudget) ||
                             !IsSoftMaskEntirelyTransparent(
                                 softMask,
                                 transform,
                                 resources,
                                 surfaceWidth,
                                 surfaceHeight,
                                 type3PaintChannelCache,
                                 activeType3PaintChannelStreams,
                                 pageContentBudget,
                                 type3GlyphBudget,
                                 depth + 1),
                         visibleShadingVisitor: name => {
                             if (requireImageMask ||
                                 !directShadings.TryGetValue(name, out PdfPageShadingResource shading) ||
                                 !shading.SupportsExactType3Projection) {
                                 supported = false;
                             }
                         },
                         invalidPatternSelectionVisitor: () => supported = false,
                         patternSelectionVisitor: selection => {
                             if (selection.BaseColorSpace?.UsesIccApproximation == true) {
                                 AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.IccColorSpaceId, selection.Name);
                                 supported = false;
                             }
                         },
                         pageWidth: surfaceWidth,
                         textClippingBudget: textClippingBudget,
                         initialHasAuthoredRenderingIntent: initialHasAuthoredRenderingIntent,
                         initialRenderingIntent: initialRenderingIntent,
                         initialFillColorSelection: initialFillColorSelection,
                         initialStrokeColorSelection: initialStrokeColorSelection,
                         outputIntentColorTransform: EffectiveOutputIntentColorTransform)) {
                PdfPageDrawingEffect invocationEffect = ResolveDrawingEffect(
                    drawingEffects,
                    invocation.PaintOrder);
                if (IsPaintSuppressedByTransparentSoftMask(
                        invocationEffect,
                        resources,
                        programState.Transform,
                        surfaceWidth,
                        surfaceHeight,
                        type3PaintChannelCache,
                        activeType3PaintChannelStreams,
                        pageContentBudget,
                        type3GlyphBudget,
                        depth + 1,
                        invocation.FillColor,
                        invocation.StrokeColor,
                        invocation.FillPattern.HasValue,
                        invocation.StrokePattern.HasValue)) {
                    continue;
                }
                if (invocation.InlineImage != null || TryGetImageXObject(resources, invocation.Name, out _, out _)) {
                    bool canProjectImage = CanProjectType3ImageInvocation(
                        invocation,
                        resources,
                        requireImageMask,
                        invocation.FillPattern,
                        diagnostics,
                        seen,
                        surfaceWidth,
                        surfaceHeight,
                        type3GlyphBudget.VisibilityGeometryBudget,
                        pageContentBudget,
                        requireInterpolation: requireIsolatedGroupSemantics);
                    if (!canProjectImage) supported = false;
                    continue;
                }
                if (!TryGetFormStream(resources, invocation.Name, out PdfStream form)) {
                    supported = false;
                    continue;
                }
                if (form.Dictionary.Items.TryGetValue("OC", out PdfObject? formOptionalContentObject) &&
                    ResolveEffectObject(formOptionalContentObject) is not PdfNull) {
                    supported = false;
                    continue;
                }
                if (ResolveEffectObject(form.Dictionary.Items.TryGetValue("Type", out PdfObject? formTypeObject) ? formTypeObject : null) is not PdfName { Name: "XObject" }) {
                    supported = false;
                    continue;
                }
                if (!TryClassifyType3TransparencyGroup(form.Dictionary, out bool isTransparencyGroup) ||
                    !TryReadFormMatrix(form.Dictionary, out Matrix2D authoredFormMatrix) ||
                    !isTransparencyGroup && !TryReadExactType3FormBox(form.Dictionary, out _)) {
                    supported = false;
                    continue;
                }
                if (ResolveXObjectPaintChannels(
                        resources,
                        invocation.Name,
                        invocation.PaintState,
                        surfaceWidth,
                        surfaceHeight,
                        type3PaintChannelCache,
                        activeType3PaintChannelStreams,
                        pageContentBudget,
                        type3GlyphBudget) == PdfType3PaintChannels.None) {
                    continue;
                }
                Matrix2D formTransform = Matrix2D.Multiply(invocation.Transform, authoredFormMatrix);
                PdfPageClipPath? formClipPath = invocation.ClipPath;
                PdfPagePatternSelection? formFillPattern = invocation.FillPattern;
                PdfPagePatternSelection? formStrokePattern = invocation.StrokePattern;
                double formSurfaceWidth = surfaceWidth;
                double formSurfaceHeight = surfaceHeight;
                if (isTransparencyGroup) {
                    if ((invocation.FillOpacity ?? 1D) <= 0D) continue;
                    if (!IsSupportedType3TransparencyGroup(form.Dictionary)) {
                        supported = false;
                        continue;
                    }
                    Type3TransparencyGroupDrawingResult boundsResult = TryGetVisibleType3TransparencyGroupBounds(
                        form.Dictionary,
                        formTransform,
                        invocation.ClipPath,
                        surfaceWidth,
                        surfaceHeight,
                        type3GlyphBudget.VisibilityGeometryBudget,
                        out PdfPageClipPath groupBounds);
                    if (boundsResult == Type3TransparencyGroupDrawingResult.Invisible) continue;
                    if (boundsResult == Type3TransparencyGroupDrawingResult.Unsupported) {
                        supported = false;
                        continue;
                    }
                    LocalizeType3TransparencyGroupProjection(
                        formTransform,
                        groupBounds,
                        surfaceHeight,
                        invocation.FillPattern,
                        invocation.StrokePattern,
                        out formTransform,
                        out formClipPath,
                        out formFillPattern,
                        out formStrokePattern);
                    formSurfaceWidth = groupBounds.Width;
                    formSurfaceHeight = groupBounds.Height;
                }
                if (!TryResolveStrictResources(form.Dictionary, resources, out PdfDictionary? resolvedFormResources) ||
                    resolvedFormResources == null) {
                    supported = false;
                    continue;
                }
                PdfDictionary formResources = resolvedFormResources;
                if (!CanProjectType3GlyphProgram(
                        form,
                        formResources,
                        invocation.PaintState.WithTransformAndClip(formTransform, formClipPath),
                        requireImageMask,
                        formFillPattern,
                        invocation.FillPatternBaseColorSpace,
                        formStrokePattern,
                        invocation.StrokePatternBaseColorSpace,
                        pageContentBudget,
                        type3GlyphBudget,
                        textClippingBudget,
                        activeStreams,
                        diagnostics,
                        seen,
                        depth + 1,
                        softMaskNestingDepth,
                        validatedSoftMaskGroups,
                        activeSoftMaskGroups,
                        activeSoftMaskForms,
                        formSurfaceWidth,
                        formSurfaceHeight,
                        requireIsolatedGroupSemantics: isTransparencyGroup,
                        initialHasAuthoredRenderingIntent: invocation.HasAuthoredRenderingIntent,
                        initialRenderingIntent: invocation.RenderingIntent,
                        initialFillColorSelection: invocation.FillColorSelection,
                        initialStrokeColorSelection: invocation.StrokeColorSelection)) supported = false;
            }
            return supported;
        } catch (Exception exception) when (IsRecoverableType3ProjectionFailure(exception)) {
            return false;
        } finally {
            activeStreams.Remove(stream);
        }
    }

    private static bool HasUsableInheritedPattern(PdfPagePatternSelection? selection) {
        if (!selection.HasValue) return false;
        if (selection.Value.ShadingPattern.HasValue) {
            return selection.Value.ShadingPattern.Value.SupportsExactType3Projection &&
                !selection.Value.BaseColorSpace.HasValue &&
                !selection.Value.Tint.HasValue &&
                selection.Value.ComponentCount == 0;
        }
        if (selection.Value.TilingPattern is not PdfPageTilingPatternResource pattern) return false;
        return !pattern.ConsumesInheritedLineState &&
            !pattern.HasMalformedStrictInvocation &&
            IsValidInheritedPatternSelection(selection.Value, pattern);
    }

    private bool CanProjectType3ImageInvocation(
        PdfPageXObjectInvocation invocation,
        PdfDictionary? resources,
        bool requireImageMask,
        PdfPagePatternSelection? inheritedFillPattern,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        double projectionPageWidth,
        double projectionPageHeight,
        VisualGeometryBudget geometryBudget,
        PageContentBudget pageContentBudget,
        bool requireInterpolation = false) {
        PdfImagePlacement placement;
        PdfDictionary imageDictionary;
        if (invocation.InlineImage != null) {
            imageDictionary = invocation.InlineImage.Stream.Dictionary;
            placement = BuildImagePlacement(
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
                invocation.PaintOrder,
                fillPattern: invocation.FillPattern,
                effectiveResources: resources,
                blendMode: invocation.BlendMode,
                hasUnsupportedBlendMode: invocation.HasUnsupportedBlendMode,
                hasSoftMask: invocation.HasSoftMask,
                hasAuthoredRenderingIntent: invocation.HasAuthoredRenderingIntent,
                renderingIntent: invocation.RenderingIntent);
        } else {
            if (!TryGetImageXObject(resources, invocation.Name, out int objectNumber, out int directStreamIdentity)) return false;
            PdfDictionary? xObjects = ResolveDictionary(resources?.Items.TryGetValue("XObject", out PdfObject? xObjectValue) == true ? xObjectValue : null);
            if (xObjects?.Items.TryGetValue(invocation.Name, out PdfObject? imageValue) != true ||
                ResolveObject(imageValue) is not PdfStream imageStream) return false;
            imageDictionary = imageStream.Dictionary;
            placement = BuildImagePlacement(
                0,
                invocation.Name,
                objectNumber,
                directStreamIdentity,
                invocation.Transform,
                invocation.ClipPath,
                invocation.FillColor,
                invocation.FillOpacity,
                paintOrder: invocation.PaintOrder,
                fillPattern: invocation.FillPattern,
                effectiveResources: resources,
                blendMode: invocation.BlendMode,
                hasUnsupportedBlendMode: invocation.HasUnsupportedBlendMode,
                hasSoftMask: invocation.HasSoftMask,
                hasAuthoredRenderingIntent: invocation.HasAuthoredRenderingIntent,
                renderingIntent: invocation.RenderingIntent);
        }

        if (IsInvisibleImagePlacement(
                placement,
                projectionPageHeight,
                projectionPageWidth,
                projectionPageHeight,
                geometryBudget)) {
            return true;
        }
        if (imageDictionary.Items.TryGetValue("OC", out PdfObject? optionalContentObject) &&
            ResolveObject(optionalContentObject) is not null and not PdfNull) return false;
        if (requireInterpolation && !ResolveType3ImageInterpolation(imageDictionary)) return false;
        if (!TryCreateImageProjection(
                placement,
                projectionPageHeight,
                projectionPageWidth,
                projectionPageHeight,
                out _,
                allowAxisAlignedFallback: false)) return false;

        IReadOnlyList<PdfExtractedImage> images;
        try {
            images = GetImagesForResources(resources, 0, new[] { placement }, colorizeImageMasks: true, pageContentBudget);
        } catch (IOException exception) when (exception is not PdfReadLimitException) {
            return false;
        } catch (NotSupportedException) {
            return false;
        }
        PdfExtractedImage? image = FindImage(images, placement);
        bool requiresOptionalCodec = RequiresOptionalImageCodec(imageDictionary.Items.TryGetValue("Filter", out PdfObject? filterObject) ? filterObject : null);
        if (requiresOptionalCodec) {
            AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.OptionalImageCodecId, invocation.Name);
        }
        CollectImageColorSpaceCapabilityDiagnostic(
            imageDictionary,
            resources,
            diagnostics,
            seen,
            invocation.Name,
            pageContentBudget,
            invocation.InlineImage?.Stream);
        if (image == null || !IsSupportedType3Image(placement, image, resources, pageContentBudget) || image.HasUnresolvedTransparencyMask || (requireImageMask && !image.IsImageMask)) return false;
        if (!requireImageMask && image.IsImageMask && inheritedFillPattern.HasValue) return false;
        if (requireImageMask && inheritedFillPattern.HasValue) {
            Type3PatternImageMaskDrawingResult result = TryPrepareInheritedPatternImageMaskDrawing(
                inheritedFillPattern,
                placement,
                image,
                projectionPageWidth,
                projectionPageHeight,
                geometryBudget,
                out _,
                out _,
                out _,
                out _,
                out _,
                out bool shadingPreparationFailed);
            if (result == Type3PatternImageMaskDrawingResult.Unsupported) {
                if (shadingPreparationFailed) {
                    AddRenderDiagnostic(
                        diagnostics,
                        seen,
                        PdfRenderCapabilities.UnsupportedShadingId,
                        inheritedFillPattern.Value.Name);
                }
                return false;
            }
        }
        return true;
    }

    private void CollectFontCapabilityDiagnostics(PdfDictionary resources, HashSet<string> invokedFonts, HashSet<string> failedType3Fonts, List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen) {
        foreach (PdfFontResource font in ResourceResolver.GetFontsForResources(resources, _objects).Values) {
            if (!invokedFonts.Contains(font.ResourceName) || font.EmbeddedTrueTypeFont != null) continue;
            string capabilityId;
            if (string.Equals(font.FontSubtype, "Type3", StringComparison.Ordinal)) {
                if (font.Type3 != null && !failedType3Fonts.Contains(font.ResourceName)) continue;
                capabilityId = PdfRenderCapabilities.Type3FontSubstitutionId;
            } else if (font.EmbeddedProgramSubtype is "Type1C" or "CIDFontType0C" or "CFF") {
                capabilityId = PdfRenderCapabilities.CffFontSubstitutionId;
            } else {
                capabilityId = PdfRenderCapabilities.FontSubstitutionId;
            }
            AddRenderDiagnostic(diagnostics, seen, capabilityId, font.ResourceName);
        }
    }

    private HashSet<string> GetUnsupportedColorSpaceResourceNames(PdfDictionary? resources, PageContentBudget pageContentBudget, HashSet<string>? invokedNames = null) {
        var unsupported = new HashSet<string>(StringComparer.Ordinal);
        if (resources == null) return unsupported;
        PdfDictionary? colorSpaces = ResolveDictionary(resources.Items.TryGetValue("ColorSpace", out PdfObject? value) ? value : null);
        if (colorSpaces == null) return unsupported;
        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (invokedNames != null && !invokedNames.Contains(entry.Key)) continue;
            if (!TryReadColorSpaceResource(
                    entry.Value,
                    pageContentBudget.TryConsumeColorFunctionEvaluation,
                    pageContentBudget.ColorFunctionResolutionContext,
                    out _)) {
                unsupported.Add(entry.Key);
            }
        }

        return unsupported;
    }

    private HashSet<string> GetApproximatedIccColorSpaceResourceNames(PdfDictionary? resources, PageContentBudget pageContentBudget, HashSet<string>? invokedNames = null) {
        var approximated = new HashSet<string>(StringComparer.Ordinal);
        if (resources == null) return approximated;
        PdfDictionary? colorSpaces = ResolveDictionary(resources.Items.TryGetValue("ColorSpace", out PdfObject? value) ? value : null);
        if (colorSpaces == null) return approximated;
        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (invokedNames != null && !invokedNames.Contains(entry.Key)) continue;
            if (TryReadColorSpaceResource(
                    entry.Value,
                    pageContentBudget.TryConsumeColorFunctionEvaluation,
                    pageContentBudget.ColorFunctionResolutionContext,
                    out PdfPageColorSpace colorSpace) && colorSpace.UsesIccApproximation) {
                approximated.Add(entry.Key);
            }
        }
        return approximated;
    }

    private void CollectShadingCapabilityDiagnostics(
        PdfDictionary resources,
        IReadOnlyCollection<string> invokedShadings,
        IReadOnlyCollection<string> invokedPatterns,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        PageContentBudget pageContentBudget) {
        PdfDictionary? shadings = ResolveDictionary(resources.Items.TryGetValue("Shading", out PdfObject? shadingValue) ? shadingValue : null);
        foreach (string name in invokedShadings) {
            if (shadings?.Items.TryGetValue(name, out PdfObject? shading) == true) {
                CollectOneShadingCapabilityDiagnostic(shading, name, diagnostics, seen, pageContentBudget);
            } else {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedShadingId, name);
            }
        }

        PdfDictionary? patterns = ResolveDictionary(resources.Items.TryGetValue("Pattern", out PdfObject? patternValue) ? patternValue : null);
        foreach (string name in invokedPatterns) {
            if (patterns?.Items.TryGetValue(name, out PdfObject? patternValueObject) != true) continue;
            PdfDictionary? pattern = ResolveDictionary(patternValueObject);
            if (TryReadInteger(pattern?.Items.TryGetValue("PatternType", out PdfObject? typeValue) == true ? typeValue : null) != 2) continue;
            if (pattern?.Items.TryGetValue("Shading", out PdfObject? shading) == true) {
                CollectOneShadingCapabilityDiagnostic(shading, name, diagnostics, seen, pageContentBudget);
            } else {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedShadingId, name);
            }
        }
    }

    private void CollectOneShadingCapabilityDiagnostic(
        PdfObject? value,
        string subject,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        PageContentBudget pageContentBudget) {
        PdfDictionary? shading = ResolveDictionary(value);
        if (shading == null || !shading.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject) ||
            !TryReadColorSpaceResource(
                colorSpaceObject,
                pageContentBudget.TryConsumeColorFunctionEvaluation,
                pageContentBudget.ColorFunctionResolutionContext,
                out PdfPageColorSpace colorSpace)) {
            AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.ColorSpaceId, subject);
        } else if (colorSpace.UsesIccApproximation) {
            AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.IccColorSpaceId, subject);
        }
        if (!TryReadShading(value, out _, pageContentBudget: pageContentBudget)) {
            AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedShadingId, subject);
        }
    }

    private void CollectPatternCapabilityDiagnostics(PdfDictionary resources, List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen) {
        PdfDictionary? patterns = ResolveDictionary(resources.Items.TryGetValue("Pattern", out PdfObject? value) ? value : null);
        if (patterns == null) return;
        foreach (KeyValuePair<string, PdfObject> entry in patterns.Items) {
            PdfObject? resolved = ResolveObject(entry.Value);
            PdfDictionary? pattern = resolved switch {
                PdfDictionary dictionary => dictionary,
                PdfStream stream => stream.Dictionary,
                _ => null
            };
            if (pattern?.Get<PdfNumber>("PatternType")?.Value == 1D) {
                if (!IsStructurallySupportedTilingPattern(resolved, pattern)) {
                    AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedTilingPatternId, entry.Key);
                }
            }
        }
    }

    private bool IsStructurallySupportedTilingPattern(PdfObject? resolved, PdfDictionary pattern) {
        int? paintType = TryReadInteger(pattern.Items.TryGetValue("PaintType", out PdfObject? paintTypeObject) ? paintTypeObject : null);
        int? tilingType = TryReadInteger(pattern.Items.TryGetValue("TilingType", out PdfObject? tilingTypeObject) ? tilingTypeObject : null);
        Matrix2D matrix = pattern.Items.TryGetValue("Matrix", out PdfObject? matrixObject)
            ? ReadPatternMatrix(matrixObject)
            : Matrix2D.Identity;
        return resolved is PdfStream &&
            (paintType == 1 || paintType == 2) &&
            tilingType >= 1 && tilingType <= 3 &&
            TryReadRectangle(pattern.Items.TryGetValue("BBox", out PdfObject? boxObject) ? boxObject : null, out (double X1, double Y1, double X2, double Y2) box) &&
            box.X2 > box.X1 && box.Y2 > box.Y1 &&
            ResolveObject(pattern.Items.TryGetValue("XStep", out PdfObject? xStepObject) ? xStepObject : null) is PdfNumber xStep &&
            ResolveObject(pattern.Items.TryGetValue("YStep", out PdfObject? yStepObject) ? yStepObject : null) is PdfNumber yStep &&
            !double.IsNaN(xStep.Value) && !double.IsInfinity(xStep.Value) && Math.Abs(xStep.Value) > 0.0000001D &&
            !double.IsNaN(yStep.Value) && !double.IsInfinity(yStep.Value) && Math.Abs(yStep.Value) > 0.0000001D &&
            IsUsableTilingPatternMatrix(matrix);
    }

    private void CollectAuxiliarySurfaceCapabilityDiagnostics(
        PdfDictionary resources,
        IReadOnlyCollection<string> invokedPatterns,
        IReadOnlyCollection<PdfStream> invokedSoftMasks,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        PdfTextClippingBudget textClippingBudget,
        int auxiliarySurfaceDepth) {
        PdfDictionary? patterns = ResolveDictionary(resources.Items.TryGetValue("Pattern", out PdfObject? patternObject) ? patternObject : null);
        foreach (string patternName in invokedPatterns) {
            if (patterns?.Items.TryGetValue(patternName, out PdfObject? patternValue) != true ||
                ResolveObject(patternValue) is not PdfStream patternStream ||
                TryReadInteger(patternStream.Dictionary.Items.TryGetValue("PatternType", out PdfObject? typeValue) ? typeValue : null) != 1) continue;
            CollectOneAuxiliarySurfaceCapabilityDiagnostics(
                patternStream,
                resources,
                diagnostics,
                seen,
                activeForms,
                pageContentBudget,
                type3GlyphBudget,
                textClippingBudget,
                auxiliarySurfaceDepth + 1,
                auxiliarySurfaceDepth);
        }

        foreach (PdfStream softMaskGroup in invokedSoftMasks) {
            CollectOneAuxiliarySurfaceCapabilityDiagnostics(
                softMaskGroup,
                resources,
                diagnostics,
                seen,
                activeForms,
                pageContentBudget,
                type3GlyphBudget,
                textClippingBudget,
                0,
                auxiliarySurfaceDepth);
        }
    }

    private void CollectOneAuxiliarySurfaceCapabilityDiagnostics(
        PdfStream stream,
        PdfDictionary parentResources,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        PdfTextClippingBudget textClippingBudget,
        int contentNestingDepth,
        int auxiliarySurfaceDepth) {
        if (!activeForms.Add(stream)) return;
        try {
            EnsureContentNestingBudget(auxiliarySurfaceDepth);
            PdfDictionary? resources = ResolveDictionary(stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourceObject) ? resourceObject : null) ?? parentResources;
            string content = PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream));
            CollectRenderCapabilityDiagnostics(content, resources, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, textClippingBudget, contentNestingDepth, auxiliarySurfaceDepth + 1);
        } finally {
            activeForms.Remove(stream);
        }
    }

    private void CollectGraphicsStateCapabilityDiagnostics(PdfDictionary resources, List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen) {
        PdfDictionary? states = ResolveDictionary(resources.Items.TryGetValue("ExtGState", out PdfObject? value) ? value : null);
        if (states == null) return;
        foreach (KeyValuePair<string, PdfObject> entry in states.Items) {
            PdfDictionary? state = ResolveDictionary(entry.Value);
            if (state == null) continue;
            if (state.Items.TryGetValue("BM", out _) && !ReadBlendMode(state).HasValue) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedBlendModeId, entry.Key);
            }
            if (state.Items.TryGetValue("SMask", out PdfObject? mask) &&
                ResolveEffectObject(mask) is not PdfName { Name: "None" } &&
                ReadSoftMask(state) == null) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedSoftMaskId, entry.Key);
            }
        }
    }

    private void CollectXObjectCapabilityDiagnostics(
        PdfDictionary resources,
        List<string> invokedXObjects,
        List<PdfPageXObjectInvocation> invokedXObjectStates,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        PdfTextClippingBudget textClippingBudget,
        int depth,
        int auxiliarySurfaceDepth) {
        PdfDictionary? xObjects = ResolveDictionary(resources.Items.TryGetValue("XObject", out PdfObject? value) ? value : null);
        if (xObjects == null) return;
        var statesByName = new Dictionary<string, List<PdfPageXObjectInvocation>>(StringComparer.Ordinal);
        for (int stateIndex = 0; stateIndex < invokedXObjectStates.Count; stateIndex++) {
            PdfPageXObjectInvocation invocation = invokedXObjectStates[stateIndex];
            if (!statesByName.TryGetValue(invocation.Name, out List<PdfPageXObjectInvocation>? states)) {
                states = new List<PdfPageXObjectInvocation>();
                statesByName.Add(invocation.Name, states);
            }
            states.Add(invocation);
        }

        var seenInvokedNames = new HashSet<string>(StringComparer.Ordinal);
        foreach (string invokedName in invokedXObjects) {
            if (!seenInvokedNames.Add(invokedName)) continue;
            if (!xObjects.Items.TryGetValue(invokedName, out PdfObject? xObject)) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.XObjectId, invokedName);
                continue;
            }

            var entry = new KeyValuePair<string, PdfObject>(invokedName, xObject);
            if (ResolveObject(entry.Value) is not PdfStream stream) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.XObjectId, entry.Key);
                continue;
            }

            string? subtype = stream.Dictionary.Get<PdfName>("Subtype")?.Name;
            if (string.Equals(subtype, "Image", StringComparison.Ordinal)) {
                CollectImageColorSpaceCapabilityDiagnostic(
                    stream.Dictionary,
                    resources,
                    diagnostics,
                    seen,
                    entry.Key,
                    pageContentBudget,
                    stream);
                if (RequiresOptionalImageCodec(stream.Dictionary.Items.TryGetValue("Filter", out PdfObject? filterObject) ? filterObject : null)) AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.OptionalImageCodecId, entry.Key);
                continue;
            }
            if (!string.Equals(subtype, "Form", StringComparison.Ordinal)) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.XObjectId, entry.Key + ":" + (subtype ?? "unknown"));
                continue;
            }

            if (!statesByName.TryGetValue(invokedName, out List<PdfPageXObjectInvocation>? states)) {
                states = new List<PdfPageXObjectInvocation>();
            }
            if (states.Count == 0) {
                states.Add(default);
            }
            var distinctStates = new HashSet<(
                string? FillName,
                bool HasFillTint,
                PdfPageColorSpace? FillBase,
                PdfPageTilingPatternResource? FillResource,
                PdfPageShadingPatternResource? FillShadingResource,
                Matrix2D FillPaintTransform,
                PdfPageColorSpace FillColorSpace,
                string? StrokeName,
                bool HasStrokeTint,
                PdfPageColorSpace? StrokeBase,
                PdfPageTilingPatternResource? StrokeResource,
                PdfPageShadingPatternResource? StrokeShadingResource,
                Matrix2D StrokePaintTransform,
                PdfPageColorSpace StrokeColorSpace,
                Matrix2D InvocationTransform,
                PdfPageClipPath? InvocationClip)>();
            for (int stateIndex = 0; stateIndex < states.Count; stateIndex++) {
                PdfPageXObjectInvocation invocation = states[stateIndex];
                var stateKey = (
                    invocation.FillPattern?.Name,
                    invocation.FillPattern?.Tint.HasValue == true,
                    invocation.FillPatternBaseColorSpace,
                    invocation.FillPattern?.TilingPattern,
                    invocation.FillPattern?.ShadingPattern,
                    invocation.FillPattern?.PaintTransform ?? Matrix2D.Identity,
                    invocation.FillColorSpace,
                    invocation.StrokePattern?.Name,
                    invocation.StrokePattern?.Tint.HasValue == true,
                    invocation.StrokePatternBaseColorSpace,
                    invocation.StrokePattern?.TilingPattern,
                    invocation.StrokePattern?.ShadingPattern,
                    invocation.StrokePattern?.PaintTransform ?? Matrix2D.Identity,
                    invocation.StrokeColorSpace,
                    invocation.Transform,
                    invocation.ClipPath);
                if (!distinctStates.Add(stateKey) || !activeForms.Add(stream)) continue;
                try {
                    PdfDictionary? formResources = ResolveDictionary(stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourceObject) ? formResourceObject : null) ?? resources;
                    Matrix2D formTransform = states.Count == 1 && string.IsNullOrEmpty(invocation.Name)
                        ? Matrix2D.Identity
                        : ApplyFormMatrix(invocation.Transform, stream.Dictionary);
                    PdfPageClipPath? formClip = CreateTransformedFormBoundingBoxClip(stream.Dictionary, formTransform);
                    if (invocation.ClipPath.HasValue) {
                        formClip = formClip.HasValue
                            ? textClippingBudget.ResolveActiveClip(invocation.ClipPath.Value, formClip.Value)
                            : invocation.ClipPath;
                    }
                    CollectRenderCapabilityDiagnostics(
                        WrapFormContentWithBoundingBoxClip(PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream)), stream.Dictionary),
                        formResources,
                        diagnostics,
                        seen,
                        activeForms,
                        pageContentBudget,
                        type3GlyphBudget,
                        textClippingBudget,
                        depth + 1,
                        auxiliarySurfaceDepth,
                        formTransform,
                        invocation.FillColorSpace,
                        invocation.FillPattern,
                        invocation.FillPatternBaseColorSpace,
                        invocation.StrokeColorSpace,
                        invocation.StrokePattern,
                        invocation.StrokePatternBaseColorSpace,
                        formClip);
                } finally {
                    activeForms.Remove(stream);
                }
            }
        }
    }

    private void CollectImageColorSpaceCapabilityDiagnostic(
        PdfDictionary image,
        PdfDictionary? resources,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        string imageName,
        PageContentBudget pageContentBudget,
        PdfStream? imageStream = null) {
        if (!image.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)) {
            return;
        }

        bool canProject = imageStream != null
            ? ResourceResolver.CanProjectImageColorSpace(
                imageStream,
                resources,
                _objects,
                _limits.MaxDecodedStreamBytes,
                EffectiveOutputIntentColorTransform,
                pageContentBudget.TryConsumeColorFunctionEvaluations,
                pageContentBudget.ColorFunctionResolutionContext)
            : ResourceResolver.CanProjectImageColorSpace(
                image,
                resources,
                _objects,
                _limits.MaxDecodedStreamBytes,
                EffectiveOutputIntentColorTransform,
                pageContentBudget.TryConsumeColorFunctionEvaluations,
                pageContentBudget.ColorFunctionResolutionContext);
        if (canProject) {
            PdfObject? diagnosticColorSpace = colorSpaceObject;
            if (ResolveObject(colorSpaceObject) is PdfName resourceName) {
                PdfDictionary? colorSpaces = ResolveDictionary(resources?.Items.TryGetValue("ColorSpace", out PdfObject? value) == true ? value : null);
                if (colorSpaces?.Items.TryGetValue(resourceName.Name, out PdfObject? resourceColorSpace) == true) diagnosticColorSpace = resourceColorSpace;
            }
            if (TryReadColorSpaceResource(
                    diagnosticColorSpace,
                    pageContentBudget.TryConsumeColorFunctionEvaluation,
                    pageContentBudget.ColorFunctionResolutionContext,
                    out PdfPageColorSpace projectedColorSpace) && projectedColorSpace.UsesIccApproximation) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.IccColorSpaceId, imageName);
            }
            return;
        }

        PdfObject? resolved = ResolveObject(colorSpaceObject);
        string subject = imageName;
        if (resolved is PdfName name) {
            subject = name.Name;
        }

        AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.ColorSpaceId, subject);
    }

    private bool RequiresOptionalImageCodec(PdfObject? value) {
        PdfObject? resolved = ResolveObject(value);
        if (resolved is PdfName name) return name.Name is "JPXDecode";
        if (resolved is not PdfArray array) return false;
        for (int i = 0; i < array.Items.Count; i++) if (RequiresOptionalImageCodec(array.Items[i])) return true;
        return false;
    }

    private PdfPageClipPath? CreateTransformedFormBoundingBoxClip(PdfDictionary dictionary, Matrix2D transform) {
        if (!TryReadBox(dictionary.Items.TryGetValue("BBox", out PdfObject? bboxObject) ? bboxObject : null, out (double X1, double Y1, double X2, double Y2) bbox) ||
            bbox.X2 <= bbox.X1 || bbox.Y2 <= bbox.Y1) return null;
        double pageHeight = GetVisualPageSize().Height;
        (double X, double Y) p0 = transform.Transform(bbox.X1, bbox.Y1);
        (double X, double Y) p1 = transform.Transform(bbox.X2, bbox.Y1);
        (double X, double Y) p2 = transform.Transform(bbox.X2, bbox.Y2);
        (double X, double Y) p3 = transform.Transform(bbox.X1, bbox.Y2);
        var commands = new[] {
            OfficePathCommand.MoveTo(p0.X, pageHeight - p0.Y),
            OfficePathCommand.LineTo(p1.X, pageHeight - p1.Y),
            OfficePathCommand.LineTo(p2.X, pageHeight - p2.Y),
            OfficePathCommand.LineTo(p3.X, pageHeight - p3.Y),
            OfficePathCommand.Close()
        };
        return PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out PdfPageClipPath clip)
            ? clip
            : null;
    }

    private void CollectAnnotationCapabilityDiagnostics(
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        PdfTextClippingBudget textClippingBudget) {
        PdfArray? annotations = ResolveArray(_pageDict.Items.TryGetValue("Annots", out PdfObject? value) ? value : null);
        if (annotations == null) return;
        EnsureAnnotationBudget(annotations);
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        for (int i = 0; i < annotations.Items.Count; i++) {
            PdfDictionary? annotation = ResolveDictionary(annotations.Items[i]);
            if (annotation == null || IsHiddenAnnotation(annotation) || HasNoVisibleAnnotationArea(annotation)) continue;
            string subtype = annotation.Get<PdfName>("Subtype")?.Name ?? "unknown";
            if (!TryGetRenderableAnnotationAppearanceStream(annotation, out PdfStream appearance, out bool synthesized)) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.AnnotationAppearanceId, subtype + "[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
                continue;
            }
            if (synthesized) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.SynthesizedAnnotationAppearanceId, subtype + "[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
            }
            if (!TryReadRectangle(annotation.Items.TryGetValue("Rect", out PdfObject? rectangleObject) ? rectangleObject : null, out (double X1, double Y1, double X2, double Y2) rectangle) ||
                !activeForms.Add(appearance)) continue;
            try {
                PdfDictionary appearanceResources = ResolveDictionary(appearance.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null)
                    ?? pageResources
                    ?? new PdfDictionary();
                string appearanceContent = WrapFormContentWithBoundingBoxClip(
                    PdfEncoding.Latin1GetString(pageContentBudget.Decode(appearance)),
                    appearance.Dictionary);
                Matrix2D appearanceTransform = Matrix2D.Multiply(
                    GetVisualPageTransform(),
                    CreateAnnotationAppearanceTransform(rectangle, appearance.Dictionary));
                CollectRenderCapabilityDiagnostics(
                    appearanceContent,
                    appearanceResources,
                    diagnostics,
                    seen,
                    activeForms,
                    pageContentBudget,
                    type3GlyphBudget,
                    textClippingBudget,
                    0,
                    1,
                    appearanceTransform,
                    initialClipPath: CreateTransformedFormBoundingBoxClip(appearance.Dictionary, appearanceTransform));
            } finally {
                activeForms.Remove(appearance);
            }
        }
    }

    private bool HasNoVisibleAnnotationArea(PdfDictionary annotation) {
        PdfArray? rectangle = ResolveArray(annotation.Items.TryGetValue("Rect", out PdfObject? value) ? value : null);
        if (rectangle == null || rectangle.Items.Count < 4 ||
            ResolveObject(rectangle.Items[0]) is not PdfNumber x1 ||
            ResolveObject(rectangle.Items[1]) is not PdfNumber y1 ||
            ResolveObject(rectangle.Items[2]) is not PdfNumber x2 ||
            ResolveObject(rectangle.Items[3]) is not PdfNumber y2) {
            return false;
        }
        return x1.Value == x2.Value || y1.Value == y2.Value;
    }

    private static void AddRenderDiagnostic(List<PdfRenderCapabilityDiagnostic> diagnostics, HashSet<string> seen, string capabilityId, string subject) {
        string key = capabilityId + "\n" + subject;
        if (seen.Add(key)) diagnostics.Add(new PdfRenderCapabilityDiagnostic(PdfRenderCapabilities.Get(capabilityId), subject));
    }
}

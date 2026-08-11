using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal IReadOnlyList<PdfRenderCapabilityDiagnostic> GetRenderCapabilityDiagnostics() {
        var diagnostics = new List<PdfRenderCapabilityDiagnostic>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var activeForms = new HashSet<PdfStream>();
        var pageContentBudget = new PageContentBudget(this);
        var type3GlyphBudget = new Type3GlyphBudget(_limits.MaxType3GlyphInvocationsPerPage);
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        CollectRenderCapabilityDiagnostics(
            GetContentStreamContent(pageContentBudget),
            resources,
            diagnostics,
            seen,
            activeForms,
            pageContentBudget,
            type3GlyphBudget,
            0,
            GetVisualPageTransform());
        CollectAnnotationCapabilityDiagnostics(diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget);
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
        int depth,
        Matrix2D? initialTransform = null,
        PdfPageColorSpace initialFillColorSpace = default,
        PdfPagePatternSelection? initialFillPattern = null,
        PdfPageColorSpace? initialFillPatternBaseColorSpace = null,
        PdfPageColorSpace initialStrokeColorSpace = default,
        PdfPagePatternSelection? initialStrokePattern = null,
        PdfPageColorSpace? initialStrokePatternBaseColorSpace = null,
        PdfPageClipPath? initialClipPath = null) {
        EnsureContentNestingBudget(depth);
        HashSet<string> unsupportedColorSpaces = GetUnsupportedColorSpaceResourceNames(resources);
        HashSet<string> approximatedIccColorSpaces = GetApproximatedIccColorSpaceResourceNames(resources);
        var invokedXObjects = new HashSet<string>(StringComparer.Ordinal);
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
                    "inline-image");
            }
        },
        maxNestingDepth: _limits.MaxContentNestingDepth,
        maxOperands: _limits.MaxContentOperands);

        if (resources == null) return;
        if (invokedShadings.Count > 0 || invokedPatterns.Count > 0) {
            _ = PdfPageContentVisualParser.Parse(
                WrapContentWithTransform(content, initialTransform ?? Matrix2D.Identity),
                GetVisualPageSize().Width,
                GetVisualPageSize().Height,
                GetGraphicsStateResources(resources),
                GetColorSpaceResources(resources),
                GetShadingResources(resources),
                GetShadingPatternResources(resources),
                tilingPatterns: null,
                GetOptionalContentVisibility(resources),
                initialFillColorSpace: initialFillColorSpace,
                initialStrokeColorSpace: initialStrokeColorSpace,
                maxOperations: _limits.MaxContentOperations,
                patternBaseColorSpaces: GetPatternBaseColorSpaceResources(resources),
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                primitiveVisitor: static _ => { },
                retainPrimitiveData: false,
                unsupportedShadingTransformVisitor: () => AddRenderDiagnostic(
                    diagnostics,
                    seen,
                    PdfRenderCapabilities.UnsupportedShadingId,
                    "transformed-radial-shading"));
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
            invokedFonts,
            invokedPatterns,
            invokedSoftMasks,
            invokedXObjectStates,
            diagnostics,
            seen);
        CollectFontCapabilityDiagnostics(resources, invokedFonts, failedType3Fonts, diagnostics, seen);
        CollectShadingCapabilityDiagnostics(resources, invokedShadings, invokedPatterns, diagnostics, seen);
        CollectPatternCapabilityDiagnostics(resources, diagnostics, seen);
        CollectGraphicsStateCapabilityDiagnostics(resources, diagnostics, seen);
        CollectXObjectCapabilityDiagnostics(resources, invokedXObjects, invokedXObjectStates, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, depth);
        CollectAuxiliarySurfaceCapabilityDiagnostics(resources, invokedPatterns, invokedSoftMasks, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, depth);
    }

    private static string? GetOperatorCapabilityId(string op) {
        switch (op) {
            case "M": return PdfRenderCapabilities.MiterLimitId;
            case "ri": return PdfRenderCapabilities.RenderingIntentId;
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
        HashSet<string> invokedFonts,
        HashSet<string> invokedPatterns,
        HashSet<PdfStream> invokedSoftMasks,
        List<PdfPageXObjectInvocation> invokedXObjectStates,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen) {
        var failures = new HashSet<string>(StringComparer.Ordinal);
        var activeStreams = new HashSet<PdfStream>();
        Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
        Dictionary<string, Func<byte[], double>> widthProviders = ResourceResolver.GetFontWidthProvidersForResources(resources, _objects);
        Dictionary<string, PdfPageColorSpace> colorSpaces = GetColorSpaceResources(resources);
        Dictionary<string, PdfPageColorSpace> patternBaseColorSpaces = GetPatternBaseColorSpaceResources(resources);
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
            graphicsStateVisitor: state => {
                if (state.SoftMask?.Group is PdfStream group) invokedSoftMasks.Add(group);
            },
            patternBaseColorSpaces: patternBaseColorSpaces,
            initialFillPattern: initialFillPattern,
            initialFillPatternBaseColorSpace: initialFillPatternBaseColorSpace,
            initialStrokePattern: initialStrokePattern,
            initialStrokePatternBaseColorSpace: initialStrokePatternBaseColorSpace);
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
            textOutputBudget: CreateTextOutputBudget(),
            pageContentBudget: new PageContentBudget(this),
            type3GlyphBudget: new Type3GlyphBudget(_limits.MaxType3GlyphInvocationsPerPage),
            requireSupportedType3Content: false);
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
                            Matrix2D.Multiply(glyph.Transform, type3.FontMatrix),
                            glyph.ClipPath,
                            type3.IsUncolored,
                            glyph.FillPattern,
                            glyph.FillPatternBaseColorSpace,
                            glyph.StrokePattern,
                            glyph.StrokePatternBaseColorSpace,
                            pageContentBudget,
                            type3GlyphBudget,
                            activeStreams,
                            diagnostics,
                            seen,
                            0)) {
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
            graphicsStateVisitor: state => {
                if (state.SoftMask?.Group is PdfStream group) invokedSoftMasks.Add(group);
            },
            patternBaseColorSpaces: patternBaseColorSpaces,
            initialFillPattern: initialFillPattern,
            initialFillPatternBaseColorSpace: initialFillPatternBaseColorSpace,
            initialStrokePattern: initialStrokePattern,
            initialStrokePatternBaseColorSpace: initialStrokePatternBaseColorSpace,
            tilingPatterns: tilingPatterns,
            shadingPatterns: GetShadingPatternResources(resources));
        invokedXObjectStates.AddRange(resolvedXObjects);
        return failures;
    }

    private bool CanProjectType3GlyphProgram(
        PdfStream stream,
        PdfDictionary resources,
        Matrix2D programTransform,
        PdfPageClipPath? programClipPath,
        bool requireImageMask,
        PdfPagePatternSelection? initialFillPattern,
        PdfPageColorSpace? initialFillPatternBaseColorSpace,
        PdfPagePatternSelection? initialStrokePattern,
        PdfPageColorSpace? initialStrokePatternBaseColorSpace,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        HashSet<PdfStream> activeStreams,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        int depth,
        double? projectionPageWidth = null,
        double? projectionPageHeight = null) {
        EnsureContentNestingBudget(depth);
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
            var validatedSoftMaskGroups = new HashSet<PdfStream>();
            var softMaskValidationBudget = new PageContentBudget(this);
            Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
            Dictionary<string, Func<byte[], double>> widthProviders = ResourceResolver.GetFontWidthProvidersForResources(resources, _objects);
            Dictionary<string, PdfPageColorSpace> colorSpaces = GetColorSpaceResources(resources);
            Dictionary<string, PdfPageColorSpace> patternBaseColorSpaces = GetPatternBaseColorSpaceResources(resources);
            var invokedPatternNames = new HashSet<string>(StringComparer.Ordinal);
            var type3PaintChannelCache = new Type3PaintChannelCache();
            var activeType3PaintChannelStreams = new HashSet<PdfStream>();
            _ = PdfPageXObjectInvocationParser.Parse(
                content,
                programTransform,
                surfaceHeight,
                GetGraphicsStateResources(resources),
                colorSpaces,
                GetOptionalContentVisibility(resources),
                maxOperations: _limits.MaxContentOperations,
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                initialClipPath: programClipPath,
                fonts: fonts,
                fontWidthProviders: widthProviders,
                patternInvocationVisitor: name => invokedPatternNames.Add(name),
                patternBaseColorSpaces: patternBaseColorSpaces,
                initialFillPattern: initialFillPattern,
                initialFillPatternBaseColorSpace: initialFillPatternBaseColorSpace,
                initialStrokePattern: initialStrokePattern,
                initialStrokePatternBaseColorSpace: initialStrokePatternBaseColorSpace,
                type3PaintChannelResolver: (font, bytes) => ResolveType3PaintChannels(
                    font,
                    bytes,
                    type3PaintChannelCache,
                    activeType3PaintChannelStreams,
                    pageContentBudget),
                xObjectPaintChannelResolver: (name, transform, clipPath, fillOpacity) => ResolveXObjectPaintChannels(
                    resources,
                    name,
                    transform,
                    clipPath,
                    fillOpacity,
                    surfaceWidth,
                    surfaceHeight,
                    type3PaintChannelCache,
                    activeType3PaintChannelStreams,
                    pageContentBudget));
            Dictionary<string, PdfPageTilingPatternResource> tilingPatterns = GetTilingPatternResources(
                resources,
                invokedPatternNames,
                textOutputBudget: CreateTextOutputBudget(),
                pageContentBudget: pageContentBudget,
                type3GlyphBudget: type3GlyphBudget,
                requireSupportedType3Content: true);
            Dictionary<string, PdfPageShadingPatternResource> shadingPatterns = GetShadingPatternResources(resources);
            bool usesFillPaint = false;
            bool usesStrokePaint = false;
            bool usesUnsupportedInheritedShadingStroke = false;
            bool usesUnsupportedInheritedShadingPlacement = false;
            _ = PdfPageContentVisualParser.Parse(
                WrapContentWithTransform(content, programTransform),
                surfaceWidth,
                surfaceHeight,
                GetGraphicsStateResources(resources),
                colorSpaces,
                GetShadingResources(resources),
                shadingPatterns,
                tilingPatterns,
                GetOptionalContentVisibility(resources),
                maxOperations: _limits.MaxContentOperations,
                patternBaseColorSpaces: patternBaseColorSpaces,
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                initialClipPath: programClipPath,
                primitiveVisitor: primitive => {
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
                    : null);
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
                         programTransform,
                         surfaceHeight,
                         GetGraphicsStateResources(resources),
                         colorSpaces,
                         GetOptionalContentVisibility(resources),
                         maxOperations: _limits.MaxContentOperations,
                         maxNestingDepth: _limits.MaxContentNestingDepth,
                         maxOperands: _limits.MaxContentOperands,
                         initialClipPath: programClipPath,
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
                                         Matrix2D.Multiply(glyph.Transform, nestedType3.FontMatrix),
                                         glyph.ClipPath,
                                         requireImageMask || nestedType3.IsUncolored,
                                         glyph.FillPattern,
                                         glyph.FillPatternBaseColorSpace,
                                         glyph.StrokePattern,
                                         glyph.StrokePatternBaseColorSpace,
                                         pageContentBudget,
                                         type3GlyphBudget,
                                         activeStreams,
                                         diagnostics,
                                         seen,
                                         depth + 1,
                                         surfaceWidth,
                                         surfaceHeight)) {
                                     supported = false;
                                 }
                             }
                             return supported;
                         },
                         type3GlyphBudgetConsumer: type3GlyphBudget.Consume,
                         unsupportedTextVisitor: () => supported = false,
                         unsupportedGraphicsEffectVisitor: () => supported = false,
                         allowSupportedGraphicsEffects: true,
                         graphicsStateVisitor: resource => {
                             if (!CanDecodeType3SoftMask(resource.SoftMask, softMaskValidationBudget, validatedSoftMaskGroups)) {
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
                                         seen);
                                 } else if (!requireImageMask) {
                                     int failureVersion = type3GlyphBudget.FailureVersion;
                                     canProject = GetTilingPatternResources(
                                             resources,
                                             new HashSet<string>(StringComparer.Ordinal) { name },
                                             textOutputBudget: CreateTextOutputBudget(),
                                             pageContentBudget: pageContentBudget,
                                             type3GlyphBudget: type3GlyphBudget,
                                             requireSupportedType3Content: true)
                                         .ContainsKey(name) &&
                                         type3GlyphBudget.FailureVersion == failureVersion;
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
                         type3PaintChannelResolver: (font, bytes) => ResolveType3PaintChannels(
                             font,
                             bytes,
                             type3PaintChannelCache,
                             activeType3PaintChannelStreams,
                             pageContentBudget),
                         xObjectPaintChannelResolver: (name, transform, clipPath, fillOpacity) => ResolveXObjectPaintChannels(
                             resources,
                             name,
                             transform,
                             clipPath,
                             fillOpacity,
                             surfaceWidth,
                             surfaceHeight,
                             type3PaintChannelCache,
                             activeType3PaintChannelStreams,
                             pageContentBudget))) {
                if (invocation.InlineImage != null || TryGetImageXObject(resources, invocation.Name, out _, out _)) {
                    bool canProjectImage = CanProjectType3ImageInvocation(
                        invocation,
                        resources,
                        requireImageMask,
                        initialFillPattern,
                        diagnostics,
                        seen,
                        surfaceWidth,
                        surfaceHeight);
                    if ((!requireImageMask && initialFillPattern.HasValue && !HasUsableInheritedPattern(initialFillPattern)) ||
                        !canProjectImage) supported = false;
                    continue;
                }
                if (!TryGetFormStream(resources, invocation.Name, out PdfStream form)) {
                    supported = false;
                    continue;
                }
                if (ResolveXObjectPaintChannels(
                        resources,
                        invocation.Name,
                        invocation.Transform,
                        invocation.ClipPath,
                        invocation.FillOpacity,
                        surfaceWidth,
                        surfaceHeight,
                        type3PaintChannelCache,
                        activeType3PaintChannelStreams,
                        pageContentBudget) == PdfType3PaintChannels.None) {
                    continue;
                }
                Matrix2D formTransform = ApplyFormMatrix(invocation.Transform, form.Dictionary);
                PdfPageClipPath? formClipPath = invocation.ClipPath;
                PdfPagePatternSelection? formFillPattern = invocation.FillPattern;
                PdfPagePatternSelection? formStrokePattern = invocation.StrokePattern;
                double formSurfaceWidth = surfaceWidth;
                double formSurfaceHeight = surfaceHeight;
                if (form.Dictionary.Items.ContainsKey("Group")) {
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
                PdfDictionary formResources = ResolveDictionary(form.Dictionary.Items.TryGetValue("Resources", out PdfObject? value) ? value : null) ?? resources;
                if (!CanProjectType3GlyphProgram(
                        form,
                        formResources,
                        formTransform,
                        formClipPath,
                        requireImageMask,
                        formFillPattern,
                        invocation.FillPatternBaseColorSpace,
                        formStrokePattern,
                        invocation.StrokePatternBaseColorSpace,
                        pageContentBudget,
                        type3GlyphBudget,
                        activeStreams,
                        diagnostics,
                        seen,
                        depth + 1,
                        formSurfaceWidth,
                        formSurfaceHeight)) supported = false;
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
        if (selection.Value.ShadingPattern.HasValue) return selection.Value.ShadingPattern.Value.SupportsExactType3Projection;
        if (selection.Value.TilingPattern is not PdfPageTilingPatternResource pattern) return false;
        return !pattern.Uncolored || selection.Value.Tint.HasValue;
    }

    private bool CanProjectType3ImageInvocation(
        PdfPageXObjectInvocation invocation,
        PdfDictionary resources,
        bool requireImageMask,
        PdfPagePatternSelection? inheritedFillPattern,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        double projectionPageWidth,
        double projectionPageHeight) {
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
                invocation.PaintOrder);
        } else {
            if (!TryGetImageXObject(resources, invocation.Name, out int objectNumber, out int directStreamIdentity)) return false;
            PdfDictionary? xObjects = ResolveDictionary(resources.Items.TryGetValue("XObject", out PdfObject? xObjectValue) ? xObjectValue : null);
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
                paintOrder: invocation.PaintOrder);
        }

        if (requireImageMask) {
            if (IsInvisibleImagePlacement(placement, projectionPageHeight, projectionPageWidth, projectionPageHeight)) {
                return true;
            }
        }

        IReadOnlyList<PdfExtractedImage> images;
        try {
            images = GetImagesForResources(resources, 0, new[] { placement }, colorizeImageMasks: true);
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
        CollectImageColorSpaceCapabilityDiagnostic(imageDictionary, resources, diagnostics, seen, invocation.Name);
        if (image == null || !image.IsImageFile || (requireImageMask && !image.IsImageMask)) return false;
        if (requireImageMask && inheritedFillPattern.HasValue) {
            Type3PatternImageMaskDrawingResult result = TryPrepareInheritedPatternImageMaskDrawing(
                inheritedFillPattern,
                placement,
                image,
                projectionPageWidth,
                projectionPageHeight,
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

    private HashSet<string> GetUnsupportedColorSpaceResourceNames(PdfDictionary? resources) {
        var unsupported = new HashSet<string>(StringComparer.Ordinal);
        if (resources == null) return unsupported;
        PdfDictionary? colorSpaces = ResolveDictionary(resources.Items.TryGetValue("ColorSpace", out PdfObject? value) ? value : null);
        if (colorSpaces == null) return unsupported;
        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (!TryReadColorSpaceResource(entry.Value, out _)) {
                unsupported.Add(entry.Key);
            }
        }

        return unsupported;
    }

    private HashSet<string> GetApproximatedIccColorSpaceResourceNames(PdfDictionary? resources) {
        var approximated = new HashSet<string>(StringComparer.Ordinal);
        if (resources == null) return approximated;
        PdfDictionary? colorSpaces = ResolveDictionary(resources.Items.TryGetValue("ColorSpace", out PdfObject? value) ? value : null);
        if (colorSpaces == null) return approximated;
        foreach (KeyValuePair<string, PdfObject> entry in colorSpaces.Items) {
            if (TryReadColorSpaceResource(entry.Value, out PdfPageColorSpace colorSpace) && colorSpace.UsesIccApproximation) {
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
        HashSet<string> seen) {
        PdfDictionary? shadings = ResolveDictionary(resources.Items.TryGetValue("Shading", out PdfObject? shadingValue) ? shadingValue : null);
        foreach (string name in invokedShadings) {
            if (shadings?.Items.TryGetValue(name, out PdfObject? shading) == true) {
                CollectOneShadingCapabilityDiagnostic(shading, name, diagnostics, seen);
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
                CollectOneShadingCapabilityDiagnostic(shading, name, diagnostics, seen);
            } else {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.UnsupportedShadingId, name);
            }
        }
    }

    private void CollectOneShadingCapabilityDiagnostic(
        PdfObject? value,
        string subject,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen) {
        PdfDictionary? shading = ResolveDictionary(value);
        if (shading == null || !shading.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject) ||
            !TryReadColorSpaceResource(colorSpaceObject, out PdfPageColorSpace colorSpace)) {
            AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.ColorSpaceId, subject);
        } else if (colorSpace.UsesIccApproximation) {
            AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.IccColorSpaceId, subject);
        }
        if (!TryReadShading(value, out _)) {
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
        int depth) {
        PdfDictionary? patterns = ResolveDictionary(resources.Items.TryGetValue("Pattern", out PdfObject? patternObject) ? patternObject : null);
        foreach (string patternName in invokedPatterns) {
            if (patterns?.Items.TryGetValue(patternName, out PdfObject? patternValue) != true ||
                ResolveObject(patternValue) is not PdfStream patternStream ||
                TryReadInteger(patternStream.Dictionary.Items.TryGetValue("PatternType", out PdfObject? typeValue) ? typeValue : null) != 1) continue;
            CollectOneAuxiliarySurfaceCapabilityDiagnostics(patternStream, resources, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, depth);
        }

        foreach (PdfStream softMaskGroup in invokedSoftMasks) {
            CollectOneAuxiliarySurfaceCapabilityDiagnostics(softMaskGroup, resources, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, depth);
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
        int depth) {
        if (!activeForms.Add(stream)) return;
        try {
            PdfDictionary? resources = ResolveDictionary(stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourceObject) ? resourceObject : null) ?? parentResources;
            string content = PdfEncoding.Latin1GetString(pageContentBudget.Decode(stream));
            CollectRenderCapabilityDiagnostics(content, resources, diagnostics, seen, activeForms, pageContentBudget, type3GlyphBudget, depth + 1);
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
        HashSet<string> invokedXObjects,
        IReadOnlyList<PdfPageXObjectInvocation> invokedXObjectStates,
        List<PdfRenderCapabilityDiagnostic> diagnostics,
        HashSet<string> seen,
        HashSet<PdfStream> activeForms,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        PdfDictionary? xObjects = ResolveDictionary(resources.Items.TryGetValue("XObject", out PdfObject? value) ? value : null);
        if (xObjects == null) return;
        foreach (string invokedName in invokedXObjects) {
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
                    entry.Key);
                if (RequiresOptionalImageCodec(stream.Dictionary.Items.TryGetValue("Filter", out PdfObject? filterObject) ? filterObject : null)) AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.OptionalImageCodecId, entry.Key);
                continue;
            }
            if (!string.Equals(subtype, "Form", StringComparison.Ordinal)) {
                AddRenderDiagnostic(diagnostics, seen, PdfRenderCapabilities.XObjectId, entry.Key + ":" + (subtype ?? "unknown"));
                continue;
            }

            List<PdfPageXObjectInvocation> states = invokedXObjectStates
                .Where(invocation => string.Equals(invocation.Name, invokedName, StringComparison.Ordinal))
                .ToList();
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
                            ? PdfPageClipPath.ResolveActiveClip(invocation.ClipPath.Value, formClip.Value)
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
                        depth + 1,
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
        string imageName) {
        if (!image.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject)) {
            return;
        }

        if (ResourceResolver.CanProjectImageColorSpace(image, resources, _objects)) {
            PdfObject? diagnosticColorSpace = colorSpaceObject;
            if (ResolveObject(colorSpaceObject) is PdfName resourceName) {
                PdfDictionary? colorSpaces = ResolveDictionary(resources?.Items.TryGetValue("ColorSpace", out PdfObject? value) == true ? value : null);
                if (colorSpaces?.Items.TryGetValue(resourceName.Name, out PdfObject? resourceColorSpace) == true) diagnosticColorSpace = resourceColorSpace;
            }
            if (TryReadColorSpaceResource(diagnosticColorSpace, out PdfPageColorSpace projectedColorSpace) && projectedColorSpace.UsesIccApproximation) {
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
        Type3GlyphBudget type3GlyphBudget) {
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

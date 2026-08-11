using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private OfficeBlendMode? ReadBlendMode(PdfDictionary state) {
        if (!state.Items.TryGetValue("BM", out PdfObject? value)) return null;
        PdfObject? resolved = ResolveEffectObject(value);
        if (resolved is PdfArray array) {
            for (int index = 0; index < array.Items.Count; index++) {
                OfficeBlendMode? candidate = MapBlendMode(ResolveEffectObject(array.Items[index]) as PdfName);
                if (candidate.HasValue) return candidate;
            }
            return null;
        }
        return MapBlendMode(resolved as PdfName);
    }

    private static OfficeBlendMode? MapBlendMode(PdfName? name) {
        switch (name?.Name) {
            case "Normal": case "Compatible": return OfficeBlendMode.Normal;
            case "Multiply": return OfficeBlendMode.Multiply;
            case "Screen": return OfficeBlendMode.Screen;
            case "Overlay": return OfficeBlendMode.Overlay;
            case "Darken": return OfficeBlendMode.Darken;
            case "Lighten": return OfficeBlendMode.Lighten;
            case "ColorDodge": return OfficeBlendMode.ColorDodge;
            case "ColorBurn": return OfficeBlendMode.ColorBurn;
            case "HardLight": return OfficeBlendMode.HardLight;
            case "SoftLight": return OfficeBlendMode.SoftLight;
            case "Difference": return OfficeBlendMode.Difference;
            case "Exclusion": return OfficeBlendMode.Exclusion;
            case "Hue": return OfficeBlendMode.Hue;
            case "Saturation": return OfficeBlendMode.Saturation;
            case "Color": return OfficeBlendMode.Color;
            case "Luminosity": return OfficeBlendMode.Luminosity;
            default: return null;
        }
    }

    private PdfPageSoftMaskResource? ReadSoftMask(PdfDictionary state) {
        PdfObject? resolved = ResolveEffectObject(state.Items.TryGetValue("SMask", out PdfObject? value) ? value : null);
        if (resolved is PdfName { Name: "None" } || resolved is not PdfDictionary mask ||
            ResolveEffectObject(mask.Items.TryGetValue("G", out PdfObject? groupObject) ? groupObject : null) is not PdfStream group) return null;
        PdfDictionary? transparency = ResolveEffectObject(
            group.Dictionary.Items.TryGetValue("Group", out PdfObject? transparencyObject) ? transparencyObject : null) as PdfDictionary;
        if (ResolveEffectObject(group.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null) is not PdfName { Name: "Form" } ||
            transparency == null ||
            ResolveEffectObject(transparency.Items.TryGetValue("S", out PdfObject? groupSubtypeObject) ? groupSubtypeObject : null) is not PdfName { Name: "Transparency" } ||
            ResolveEffectObject(mask.Items.TryGetValue("S", out PdfObject? modeObject) ? modeObject : null) is not PdfName modeName ||
            (modeName.Name != "Alpha" && modeName.Name != "Luminosity")) return null;
        if (mask.Items.TryGetValue("TR", out PdfObject? transferObject) &&
            ResolveEffectObject(transferObject) is not PdfName { Name: "Identity" }) return null;
        if (transparency.Items.TryGetValue("I", out PdfObject? isolatedObject) &&
            ResolveEffectObject(isolatedObject) is not PdfBoolean) return null;
        if (transparency.Items.TryGetValue("K", out PdfObject? knockoutObject) &&
            ResolveEffectObject(knockoutObject) is not PdfBoolean { Value: false }) return null;
        PdfName? groupColorSpace = null;
        if (transparency.Items.TryGetValue("CS", out PdfObject? colorSpaceObject)) {
            groupColorSpace = ResolveEffectObject(colorSpaceObject) as PdfName;
            if (groupColorSpace?.Name != "DeviceGray" && groupColorSpace?.Name != "DeviceRGB") return null;
        }
        OfficeSoftMaskMode mode = modeName.Name == "Luminosity" ? OfficeSoftMaskMode.Luminosity : OfficeSoftMaskMode.Alpha;
        OfficeColor backdrop = OfficeColor.Transparent;
        if (mode == OfficeSoftMaskMode.Luminosity &&
            mask.Items.TryGetValue("BC", out PdfObject? backdropObject)) {
            if (ResolveEffectObject(backdropObject) is not PdfArray components) return null;
            IReadOnlyList<double> values = ReadNumberArray(components);
            int expectedComponents = groupColorSpace?.Name == "DeviceGray"
                ? 1
                : groupColorSpace?.Name == "DeviceRGB" ? 3 : values.Count;
            if ((expectedComponents != 1 && expectedComponents != 3) || values.Count != expectedComponents) return null;
            backdrop = values.Count == 1
                ? OfficeColor.FromRgb(ToColorByte(values[0]), ToColorByte(values[0]), ToColorByte(values[0]))
                : OfficeColor.FromRgb(ToColorByte(values[0]), ToColorByte(values[1]), ToColorByte(values[2]));
        }
        return new PdfPageSoftMaskResource(group, mode, backdrop);
    }

    private PdfObject? ResolveEffectObject(PdfObject? value) {
        PdfObject? current = value;
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        int maximumDepth = Math.Max(1, _limits.MaxContentNestingDepth);
        for (int depth = 0; depth < maximumDepth && current is PdfReference reference; depth++) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject indirect)) {
                return null;
            }
            current = indirect.Value;
        }
        return current is PdfReference ? null : current;
    }

    private bool CanDecodeType3SoftMask(
        PdfPageSoftMaskResource? resource,
        Matrix2D groupTransform,
        PageContentBudget pageContentBudget,
        Dictionary<(PdfStream Group, Matrix2D Transform), int> validatedGroups,
        Type3GlyphBudget type3GlyphBudget,
        int contentNestingDepth,
        HashSet<PdfStream>? activeType3Glyphs = null,
        SoftMaskNestingDepth? nestingDepth = null) {
        nestingDepth ??= new SoftMaskNestingDepth(contentNestingDepth);
        return CanDecodeType3SoftMask(
            resource,
            groupTransform,
            pageContentBudget,
            validatedGroups,
            type3GlyphBudget,
            new HashSet<PdfStream>(),
            new HashSet<PdfStream>(),
            activeType3Glyphs ?? new HashSet<PdfStream>(),
            contentNestingDepth,
            nestingDepth);
    }

    private bool CanDecodeType3SoftMask(
        PdfPageSoftMaskResource? resource,
        Matrix2D groupTransform,
        PageContentBudget pageContentBudget,
        Dictionary<(PdfStream Group, Matrix2D Transform), int> validatedGroups,
        Type3GlyphBudget type3GlyphBudget,
        HashSet<PdfStream> activeGroups,
        HashSet<PdfStream> activeForms,
        HashSet<PdfStream> activeType3Glyphs,
        int contentNestingDepth,
        SoftMaskNestingDepth nestingDepth) {
        if (resource == null) return true;
        EnsureContentNestingBudget(contentNestingDepth);
        nestingDepth.Maximum = Math.Max(nestingDepth.Maximum, contentNestingDepth);
        var cacheKey = (resource.Group, groupTransform);
        if (validatedGroups.TryGetValue(cacheKey, out int cachedNestingSpan)) {
            int cachedMaximumDepth = contentNestingDepth + cachedNestingSpan;
            EnsureContentNestingBudget(cachedMaximumDepth);
            nestingDepth.Maximum = Math.Max(nestingDepth.Maximum, cachedMaximumDepth);
            return true;
        }
        if (!activeGroups.Add(resource.Group)) return false;
        try {
            if (Filters.StreamDecoder.GetUnsupportedFilters(resource.Group.Dictionary, _objects).Count != 0) return false;
            string content = WrapFormContentWithBoundingBoxClip(
                PdfEncoding.Latin1GetString(pageContentBudget.Decode(resource.Group)),
                resource.Group.Dictionary);
            PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
            PdfDictionary? resources = ResolveDictionary(
                resource.Group.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourceObject)
                    ? resourceObject
                    : null) ?? pageResources;
            var groupNestingDepth = new SoftMaskNestingDepth(contentNestingDepth);
            bool supported = CanDecodeType3SoftMasksInContent(
                content,
                resources,
                groupTransform,
                pageContentBudget,
                validatedGroups,
                type3GlyphBudget,
                activeGroups,
                activeForms,
                activeType3Glyphs,
                contentNestingDepth,
                groupNestingDepth);
            nestingDepth.Maximum = Math.Max(nestingDepth.Maximum, groupNestingDepth.Maximum);
            if (supported) validatedGroups[cacheKey] = groupNestingDepth.Maximum - contentNestingDepth;
            return supported;
        } catch (PdfReadLimitException) {
            throw;
        } catch (IOException) {
            return false;
        } catch (InvalidDataException) {
            return false;
        } catch (NotSupportedException) {
            return false;
        } finally {
            activeGroups.Remove(resource.Group);
        }
    }

    private bool CanDecodeType3SoftMasksInContent(
        string content,
        PdfDictionary? resources,
        Matrix2D baseTransform,
        PageContentBudget pageContentBudget,
        Dictionary<(PdfStream Group, Matrix2D Transform), int> validatedGroups,
        Type3GlyphBudget type3GlyphBudget,
        HashSet<PdfStream> activeGroups,
        HashSet<PdfStream> activeForms,
        HashSet<PdfStream> activeType3Glyphs,
        int contentNestingDepth,
        SoftMaskNestingDepth nestingDepth,
        PdfPageXObjectInvocation? initialState = null) {
        EnsureContentNestingBudget(contentNestingDepth);
        nestingDepth.Maximum = Math.Max(nestingDepth.Maximum, contentNestingDepth);
        bool supported = true;
        Dictionary<string, PdfFontResource> fonts = ResourceResolver.GetFontsForResources(resources, _objects);
        Dictionary<string, Func<byte[], double>> widthProviders = resources == null
            ? new Dictionary<string, Func<byte[], double>>(StringComparer.Ordinal)
            : ResourceResolver.GetFontWidthProvidersForResources(resources, _objects);
        (double Width, double Height) visualPageSize = GetVisualPageSize();
        Dictionary<string, PdfPageColorSpace> colorSpaces = GetColorSpaceResources(resources);
        Dictionary<string, PdfPageColorSpace> patternBaseColorSpaces = GetPatternBaseColorSpaceResources(resources);
        var type3PaintChannelCache = new Dictionary<PdfStream, PdfType3PaintChannels>();
        var activeType3PaintChannelStreams = new HashSet<PdfStream>();
        var invokedPatternNames = new HashSet<string>(StringComparer.Ordinal);
        _ = PdfPageXObjectInvocationParser.Parse(
            content,
            baseTransform,
            visualPageSize.Height,
            GetGraphicsStateResources(resources),
            colorSpaces,
            GetOptionalContentVisibility(resources),
            initialFillColor: initialState?.FillColor,
            initialFillColorSpace: initialState?.FillColorSpace ?? default,
            initialFillOpacity: initialState?.FillOpacity,
            initialClipPath: initialState?.ClipPath,
            initialStrokeColor: initialState?.StrokeColor,
            initialStrokeColorSpace: initialState?.StrokeColorSpace ?? default,
            initialStrokeOpacity: initialState?.StrokeOpacity,
            initialStrokeWidth: initialState?.StrokeWidth,
            initialStrokeDashStyle: initialState?.StrokeDashStyle,
            initialStrokeLineCap: initialState?.StrokeLineCap,
            initialStrokeLineJoin: initialState?.StrokeLineJoin,
            maxOperations: _limits.MaxContentOperations,
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            fonts: fonts,
            fontWidthProviders: widthProviders,
            patternInvocationVisitor: name => invokedPatternNames.Add(name),
            patternBaseColorSpaces: patternBaseColorSpaces,
            initialFillPattern: initialState?.FillPattern,
            initialFillPatternBaseColorSpace: initialState?.FillPatternBaseColorSpace,
            initialStrokePattern: initialState?.StrokePattern,
            initialStrokePatternBaseColorSpace: initialState?.StrokePatternBaseColorSpace,
            type3PaintChannelResolver: (font, bytes) => ResolveType3PaintChannels(
                font,
                bytes,
                type3PaintChannelCache,
                activeType3PaintChannelStreams),
            xObjectPaintChannelResolver: (name, transform, clipPath) => ResolveXObjectPaintChannels(
                resources,
                name,
                transform,
                clipPath,
                visualPageSize.Width,
                visualPageSize.Height,
                type3PaintChannelCache,
                activeType3PaintChannelStreams));
        Dictionary<string, PdfPageTilingPatternResource> tilingPatterns = GetTilingPatternResources(
            resources,
            invokedPatternNames,
            textOutputBudget: CreateTextOutputBudget(),
            pageContentBudget: pageContentBudget,
            type3GlyphBudget: type3GlyphBudget,
            requireSupportedType3Content: false);
        Dictionary<string, PdfPageShadingPatternResource> shadingPatterns = GetShadingPatternResources(resources);
        Dictionary<string, PdfPageShadingResource> shadings = GetShadingResources(resources);
        _ = PdfPageContentVisualParser.Parse(
            WrapContentWithTransform(content, baseTransform),
            visualPageSize.Width,
            visualPageSize.Height,
            GetGraphicsStateResources(resources),
            colorSpaces,
            shadings,
            shadingPatterns,
            tilingPatterns,
            GetOptionalContentVisibility(resources),
            initialClipPath: initialState?.ClipPath,
            initialFillColor: initialState?.FillColor,
            initialFillColorSpace: initialState?.FillColorSpace ?? default,
            initialFillOpacity: initialState?.FillOpacity,
            initialStrokeColor: initialState?.StrokeColor,
            initialStrokeColorSpace: initialState?.StrokeColorSpace ?? default,
            initialStrokeOpacity: initialState?.StrokeOpacity,
            initialStrokeWidth: initialState?.StrokeWidth,
            initialStrokeDashStyle: initialState?.StrokeDashStyle,
            initialStrokeLineCap: initialState?.StrokeLineCap,
            initialStrokeLineJoin: initialState?.StrokeLineJoin,
            maxOperations: _limits.MaxContentOperations,
            patternBaseColorSpaces: patternBaseColorSpaces,
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            retainPrimitiveData: false,
            unsupportedShadingTransformVisitor: () => supported = false,
            requireExactType3ShadingProjection: true,
            authoredShadingInvocationVisitor: name => {
                if (!shadings.TryGetValue(name, out PdfPageShadingResource shading) ||
                    !shading.SupportsExactType3Projection) {
                    supported = false;
                }
            });
        if (!supported) return false;
        var validationDiagnostics = new List<PdfRenderCapabilityDiagnostic>();
        var validationDiagnosticKeys = new HashSet<string>(StringComparer.Ordinal);
        IReadOnlyList<PdfPageXObjectInvocation> invocations = PdfPageXObjectInvocationParser.Parse(
            content,
            baseTransform,
            visualPageSize.Height,
            GetGraphicsStateResources(resources),
            colorSpaces,
            GetOptionalContentVisibility(resources),
            initialFillColor: initialState?.FillColor,
            initialFillColorSpace: initialState?.FillColorSpace ?? default,
            initialFillOpacity: initialState?.FillOpacity,
            initialClipPath: initialState?.ClipPath,
            initialStrokeColor: initialState?.StrokeColor,
            initialStrokeColorSpace: initialState?.StrokeColorSpace ?? default,
            initialStrokeOpacity: initialState?.StrokeOpacity,
            initialStrokeWidth: initialState?.StrokeWidth,
            initialStrokeDashStyle: initialState?.StrokeDashStyle,
            initialStrokeLineCap: initialState?.StrokeLineCap,
            initialStrokeLineJoin: initialState?.StrokeLineJoin,
            maxOperations: _limits.MaxContentOperations,
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            fonts: fonts,
            fontWidthProviders: widthProviders,
            type3TextVisitor: invocation => {
                for (int glyphIndex = 0; glyphIndex < invocation.Glyphs.Count; glyphIndex++) {
                    PdfPageType3GlyphInvocation glyph = invocation.Glyphs[glyphIndex];
                    if (glyph.Font.Type3 is not PdfType3FontResource type3 ||
                        !type3.TryGetGlyph(glyph.CharacterCode, out PdfStream glyphStream) ||
                        !CanProjectType3GlyphProgram(
                            glyphStream,
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
                            activeType3Glyphs,
                            validationDiagnostics,
                            validationDiagnosticKeys,
                            contentNestingDepth + 1,
                            nestingDepth,
                            validatedGroups)) {
                        supported = false;
                    }
                }
                return supported;
            },
            type3GlyphBudgetConsumer: type3GlyphBudget.Consume,
            unsupportedGraphicsEffectVisitor: () => supported = false,
            graphicsStateVisitor: (state, stateTransform) => {
                if (!CanDecodeType3SoftMask(
                        state.SoftMask,
                        stateTransform,
                        pageContentBudget,
                        validatedGroups,
                        type3GlyphBudget,
                        activeGroups,
                        activeForms,
                        activeType3Glyphs,
                        contentNestingDepth + 1,
                        nestingDepth)) {
                    supported = false;
                }
            },
            allowSupportedGraphicsEffects: true,
            patternBaseColorSpaces: patternBaseColorSpaces,
            initialFillPattern: initialState?.FillPattern,
            initialFillPatternBaseColorSpace: initialState?.FillPatternBaseColorSpace,
            initialStrokePattern: initialState?.StrokePattern,
            initialStrokePatternBaseColorSpace: initialState?.StrokePatternBaseColorSpace,
            tilingPatterns: tilingPatterns,
            shadingPatterns: shadingPatterns,
            type3PaintChannelResolver: (font, bytes) => ResolveType3PaintChannels(
                font,
                bytes,
                type3PaintChannelCache,
                activeType3PaintChannelStreams),
            xObjectPaintChannelResolver: (name, transform, clipPath) => ResolveXObjectPaintChannels(
                resources,
                name,
                transform,
                clipPath,
                visualPageSize.Width,
                visualPageSize.Height,
                type3PaintChannelCache,
                activeType3PaintChannelStreams));
        if (!supported) return false;

        for (int index = 0; index < invocations.Count; index++) {
            PdfPageXObjectInvocation invocation = invocations[index];
            if (invocation.InlineImage != null || TryGetImageXObject(resources, invocation.Name, out _, out _)) {
                if (!CanProjectType3ImageInvocation(
                        invocation,
                        resources,
                        requireImageMask: false,
                        inheritedFillPattern: null,
                        diagnostics: validationDiagnostics,
                        seen: validationDiagnosticKeys,
                        projectionPageWidth: visualPageSize.Width,
                        projectionPageHeight: visualPageSize.Height)) {
                    return false;
                }
                continue;
            }
            if (!TryGetFormStream(resources, invocation.Name, out PdfStream form) || !activeForms.Add(form)) return false;
            try {
                if (Filters.StreamDecoder.GetUnsupportedFilters(form.Dictionary, _objects).Count != 0) return false;
                if (form.Dictionary.Items.ContainsKey("Group")) return false;
                string formContent = WrapFormContentWithBoundingBoxClip(
                    PdfEncoding.Latin1GetString(pageContentBudget.Decode(form)),
                    form.Dictionary);
                PdfDictionary? formResources = ResolveDictionary(
                    form.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourceObject)
                        ? formResourceObject
                        : null) ?? resources;
                if (!CanDecodeType3SoftMasksInContent(
                        formContent,
                        formResources,
                        ApplyFormMatrix(invocation.Transform, form.Dictionary),
                        pageContentBudget,
                        validatedGroups,
                        type3GlyphBudget,
                        activeGroups,
                        activeForms,
                        activeType3Glyphs,
                        contentNestingDepth + 1,
                        nestingDepth,
                        invocation)) {
                    return false;
                }
            } finally {
                activeForms.Remove(form);
            }
        }
        return true;
    }

    private sealed class SoftMaskNestingDepth {
        internal SoftMaskNestingDepth(int maximum) {
            Maximum = maximum;
        }

        internal int Maximum { get; set; }
    }

    private IReadOnlyList<PdfPageDrawingEffectTransition> GetGraphicsEffectTransitions(Matrix2D pageTransform, double pageHeight, PageContentBudget? pageContentBudget = null) {
        pageContentBudget ??= new PageContentBudget(this);
        var transitions = new List<PdfPageDrawingEffectTransition>();
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        string content = GetContentStreamContent(pageContentBudget);
        if (content.Length == 0) return Array.Empty<PdfPageDrawingEffectTransition>();
        var activeForms = new HashSet<PdfStream>();
        CollectGraphicsEffectTransitions(
            content,
            resources,
            pageTransform,
            pageHeight,
            transitions,
            activeForms,
            PdfPageDrawingEffect.Default,
            pageContentBudget: pageContentBudget,
            contentOrderPrefix: PdfContentOrderKey.Root);
        SortGraphicsEffectTransitions(transitions);
        return transitions.Count == 0 ? Array.Empty<PdfPageDrawingEffectTransition>() : transitions.AsReadOnly();
    }

    private void CollectGraphicsEffectTransitions(
        string content,
        PdfDictionary? resources,
        Matrix2D baseTransform,
        double pageHeight,
        List<PdfPageDrawingEffectTransition> transitions,
        HashSet<PdfStream> activeForms,
        PdfPageDrawingEffect initialEffect,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        PdfPageClipPath? initialClipPath = null,
        OfficeColor? initialFillColor = null,
        PdfPageColorSpace initialFillColorSpace = default,
        double? initialFillOpacity = null,
        OfficeColor? initialStrokeColor = null,
        PdfPageColorSpace initialStrokeColorSpace = default,
        double? initialStrokeOpacity = null,
        double? initialStrokeWidth = null,
        OfficeStrokeDashStyle? initialStrokeDashStyle = null,
        OfficeStrokeLineCap? initialStrokeLineCap = null,
        OfficeStrokeLineJoin? initialStrokeLineJoin = null,
        int contentNestingDepth = 0,
        PageContentBudget? pageContentBudget = null,
        PdfContentOrderKey? contentOrderPrefix = null,
        bool skipTransparencyGroupForms = false) {
        EnsureContentNestingBudget(contentNestingDepth);
        pageContentBudget ??= new PageContentBudget(this);
        Dictionary<string, PdfPageGraphicsStateResource> graphicsStates = GetGraphicsStateResources(resources);
        IReadOnlyList<PdfPageDrawingEffectTransition> parsed = PdfPageGraphicsEffectTimelineParser.Parse(
            content,
            graphicsStates,
            initialEffect,
            baseTransform,
            contentOrderPrefix,
            paintOrderBase,
            paintOrderScale,
            paintOrderOffset,
            _limits.MaxContentOperations,
            _limits.MaxContentNestingDepth,
            _limits.MaxContentOperands);
        var local = new List<PdfPageDrawingEffectTransition>(parsed.Count);
        for (int transitionIndex = 0; transitionIndex < parsed.Count; transitionIndex++) {
            PdfPageDrawingEffectTransition transition = parsed[transitionIndex];
            PdfPageDrawingEffect effect = transition.Effect.SoftMask != null && !transition.Effect.SoftMaskTransform.HasValue
                ? transition.Effect.WithSoftMaskTransform(baseTransform)
                : transition.Effect;
            var resolvedTransition = new PdfPageDrawingEffectTransition(
                transition.PaintOrder,
                effect,
                transition.ContentOrderKey,
                contentNestingDepth);
            local.Add(resolvedTransition);
            transitions.Add(resolvedTransition);
        }

        foreach (PdfPageXObjectInvocation invocation in PdfPageXObjectInvocationParser.Parse(
                     content,
                     baseTransform,
                     pageHeight,
                     graphicsStates,
                     GetColorSpaceResources(resources),
                     GetOptionalContentVisibility(resources),
                     initialFillColor,
                     initialFillColorSpace,
                     initialFillOpacity,
                     paintOrderBase,
                     paintOrderScale,
                     paintOrderOffset,
                     initialClipPath,
                     initialStrokeColor,
                     initialStrokeColorSpace,
                     initialStrokeOpacity,
                     initialStrokeWidth,
                     initialStrokeDashStyle,
                     initialStrokeLineCap,
                     initialStrokeLineJoin,
                     _limits.MaxContentOperations,
                     _limits.MaxContentNestingDepth,
                     _limits.MaxContentOperands)) {
            if (!TryGetFormStream(resources, invocation.Name, out PdfStream formStream) || !activeForms.Add(formStream)) continue;
            PdfContentOrderKey? formOrderPrefix = contentOrderPrefix?.Append(invocation.SourceOperatorIndex);
            PdfPageDrawingEffect inherited = ResolveDrawingEffect(local, invocation.PaintOrder, initialEffect, formOrderPrefix);
            try {
                PdfDictionary dictionary = formStream.Dictionary;
                if (skipTransparencyGroupForms && dictionary.Items.ContainsKey("Group")) {
                    continue;
                }
                PdfDictionary? formResources = ResolveDictionary(dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) ?? resources;
                Matrix2D formTransform = ApplyFormMatrix(invocation.Transform, dictionary);
                string formContent = WrapFormContentWithBoundingBoxClip(PdfEncoding.Latin1GetString(pageContentBudget.Decode(formStream)), dictionary);
                CollectGraphicsEffectTransitions(
                    formContent,
                    formResources,
                    formTransform,
                    pageHeight,
                    transitions,
                    activeForms,
                    inherited,
                    invocation.PaintOrder,
                    paintOrderScale * 0.000000001D,
                    initialClipPath: invocation.ClipPath,
                    initialFillColor: invocation.FillColor,
                    initialFillColorSpace: invocation.FillColorSpace,
                    initialFillOpacity: invocation.FillOpacity,
                    initialStrokeColor: invocation.StrokeColor,
                    initialStrokeColorSpace: invocation.StrokeColorSpace,
                    initialStrokeOpacity: invocation.StrokeOpacity,
                    initialStrokeWidth: invocation.StrokeWidth,
                    initialStrokeDashStyle: invocation.StrokeDashStyle,
                    initialStrokeLineCap: invocation.StrokeLineCap,
                    initialStrokeLineJoin: invocation.StrokeLineJoin,
                    contentNestingDepth: contentNestingDepth + 1,
                    pageContentBudget: pageContentBudget,
                    contentOrderPrefix: formOrderPrefix,
                    skipTransparencyGroupForms: skipTransparencyGroupForms);
                transitions.Add(new PdfPageDrawingEffectTransition(
                    invocation.PaintOrder + (Math.Abs(paintOrderScale) * 0.25D),
                    inherited,
                    formOrderPrefix?.Append(int.MaxValue),
                    contentNestingDepth));
            } finally {
                activeForms.Remove(formStream);
            }
        }
    }

    private static PdfPageDrawingEffect ResolveDrawingEffect(
        IReadOnlyList<PdfPageDrawingEffectTransition> transitions,
        double paintOrder,
        PdfPageDrawingEffect? initial = null,
        PdfContentOrderKey? contentOrderKey = null) {
        PdfPageDrawingEffect effect = initial ?? PdfPageDrawingEffect.Default;
        for (int i = 0; i < transitions.Count; i++) {
            PdfPageDrawingEffectTransition transition = transitions[i];
            if (contentOrderKey != null && transition.ContentOrderKey != null) {
                if (transition.ContentOrderKey.CompareTo(contentOrderKey) > 0) break;
            } else if (transition.PaintOrder > paintOrder) {
                break;
            }
            effect = transition.Effect;
        }
        return effect;
    }

    private static void SortGraphicsEffectTransitions(List<PdfPageDrawingEffectTransition> transitions) {
        transitions.Sort(static (left, right) => {
            if (left.ContentOrderKey != null && right.ContentOrderKey != null) {
                int contentOrder = left.ContentOrderKey.CompareTo(right.ContentOrderKey);
                if (contentOrder != 0) return contentOrder;
            }
            return left.PaintOrder.CompareTo(right.PaintOrder);
        });
    }

    private static void OverlayDrawingEffects(
        List<PdfPageDrawingElement> elements,
        IReadOnlyList<PdfPageDrawingEffectTransition> transitions) {
        for (int i = 0; i < elements.Count; i++) {
            PdfPageDrawingElement element = elements[i];
            PdfPageDrawingEffect inherited = ResolveDrawingEffect(
                transitions,
                element.PaintOrder,
                contentOrderKey: element.ContentOrderKey);
            elements[i] = element.WithEffect(element.Effect.OverlayOn(inherited));
        }
    }

    private OfficeDrawingSoftMask GetOrCreateSoftMask(
        PdfPageSoftMaskResource resource,
        double width,
        double height,
        Matrix2D pageTransform,
        Dictionary<(PdfStream Group, OfficeSoftMaskMode Mode, OfficeColor Backdrop, Matrix2D Transform, double Width, double Height), OfficeDrawingSoftMask> cache,
        HashSet<PdfStream> active,
        TextContentParser.TextOutputBudget textOutputBudget,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget) {
        var cacheKey = (resource.Group, resource.Mode, resource.BackdropColor, pageTransform, width, height);
        if (cache.TryGetValue(cacheKey, out OfficeDrawingSoftMask? existing)) return existing;
        if (!active.Add(resource.Group)) {
            return new OfficeDrawingSoftMask(new OfficeDrawing(width, height), resource.Mode, backdropColor: resource.BackdropColor);
        }
        try {
            OfficeDrawing drawing = CreateFormDrawing(resource.Group, width, height, pageTransform, cache, active, textOutputBudget, pageContentBudget, type3GlyphBudget);
            var mask = new OfficeDrawingSoftMask(drawing, resource.Mode, backdropColor: resource.BackdropColor);
            cache[cacheKey] = mask;
            return mask;
        } finally {
            active.Remove(resource.Group);
        }
    }

    private OfficeDrawing CreateFormDrawing(
        PdfStream form,
        double width,
        double height,
        Matrix2D pageTransform,
        Dictionary<(PdfStream Group, OfficeSoftMaskMode Mode, OfficeColor Backdrop, Matrix2D Transform, double Width, double Height), OfficeDrawingSoftMask> softMasks,
        HashSet<PdfStream> activeSoftMasks,
        TextContentParser.TextOutputBudget textOutputBudget,
        PageContentBudget pageContentBudget,
        Type3GlyphBudget type3GlyphBudget) {
        var drawing = new OfficeDrawing(width, height);
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        PdfDictionary? resources = ResolveDictionary(form.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourceObject) ? resourceObject : null) ?? pageResources;
        RegisterEmbeddedFonts(drawing, resources, new HashSet<PdfStream>(), 0);
        string content = WrapFormContentWithBoundingBoxClip(PdfEncoding.Latin1GetString(pageContentBudget.Decode(form)), form.Dictionary);
        if (content.Length == 0) return drawing;
        Matrix2D transform = ApplyFormMatrix(pageTransform, form.Dictionary);
        var activeForms = new HashSet<PdfStream>();
        var elements = new List<PdfPageDrawingElement>();
        var primitives = new List<PdfPageVisualPrimitive>();
        var renderedType3PaintOrders = new HashSet<double>();
        CollectVisualPrimitivesAndForms(
            content,
            resources,
            transform,
            width,
            height,
            primitives.Add,
            activeForms,
            renderedType3PaintOrders: renderedType3PaintOrders,
            type3GlyphBudget: type3GlyphBudget,
            allowSupportedType3TransparencyGroups: true,
            type3ImageVisitor: (placement, image, effect) => elements.Add(PdfPageDrawingElement.FromImage(placement, image, elements.Count).WithEffect(effect)),
            type3PrimitiveVisitor: (primitive, effect) => elements.Add(PdfPageDrawingElement.FromPrimitive(primitive, elements.Count).WithEffect(effect)),
            type3GroupVisitor: (group, transform, paintOrder, key, effect) => elements.Add(PdfPageDrawingElement.FromGroup(group, transform, paintOrder, key, elements.Count).WithEffect(effect)),
            textOutputBudget: textOutputBudget,
            pageContentBudget: pageContentBudget,
            contentOrderPrefix: PdfContentOrderKey.Root);
        for (int i = 0; i < primitives.Count; i++) elements.Add(PdfPageDrawingElement.FromPrimitive(primitives[i], elements.Count));

        var spans = new List<PdfTextSpan>();
        Dictionary<string, Func<byte[], int, string>> decoders = MergeDecoders(
            ResourceResolver.GetBudgetedFontDecoders(_pageDict, _objects),
            ResourceResolver.GetBudgetedFontDecodersForForm(form.Dictionary, _objects));
        Dictionary<string, Func<byte[], double>> widthProviders = MergeWidthProviders(ResourceResolver.GetFontWidthProviders(_pageDict, _objects), ResourceResolver.GetFontWidthProviders(form.Dictionary, _objects));
        Dictionary<string, PdfFontResource> fonts = MergeFonts(ResourceResolver.GetFontsForResources(pageResources, _objects), ResourceResolver.GetFontsForResources(resources, _objects));
        string transformedContent = WrapContentWithTransform(content, transform, out int transformedOffset);
        CollectTextAndForms(
            transformedContent,
            resources,
            decoders,
            widthProviders,
            fonts,
            spans,
            activeForms,
            height,
            paintOrderOffset: -transformedOffset,
            useLogicalTextFilters: false,
            textOutputBudget: textOutputBudget,
            pageContentBudget: pageContentBudget,
            contentOrderPrefix: PdfContentOrderKey.Root,
            contentOrderOffset: -transformedOffset);
        for (int i = 0; i < spans.Count; i++) {
            if (renderedType3PaintOrders.Contains(spans[i].PaintOrder)) continue;
            elements.Add(PdfPageDrawingElement.FromText(spans[i], elements.Count));
        }

        var placements = new List<PdfImagePlacement>();
        CollectImagePlacementsAndForms(content, resources, 0, transform, height, placements, activeForms, pageContentBudget: pageContentBudget);
        if (placements.Count > 0) {
            IReadOnlyList<PdfExtractedImage> images = GetImagesForResources(resources, 0, placements, colorizeImageMasks: true);
            for (int i = 0; i < placements.Count; i++) {
                PdfExtractedImage? image = FindImage(images, placements[i]);
                if (image != null) elements.Add(PdfPageDrawingElement.FromImage(placements[i], image, elements.Count));
            }
        }
        var enclosingEffects = new List<PdfPageDrawingEffectTransition>();
        CollectGraphicsEffectTransitions(
            content,
            resources,
            transform,
            height,
            enclosingEffects,
            new HashSet<PdfStream>(),
            PdfPageDrawingEffect.Default,
            pageContentBudget: pageContentBudget,
            contentOrderPrefix: PdfContentOrderKey.Root);
        SortGraphicsEffectTransitions(enclosingEffects);
        OverlayDrawingEffects(elements, enclosingEffects);

        SortDrawingElements(elements);
        for (int i = 0; i < elements.Count; i++) {
            AddDrawingElement(drawing, height, transform, elements[i], softMasks, activeSoftMasks, textOutputBudget, pageContentBudget, type3GlyphBudget);
        }
        return drawing;
    }
}

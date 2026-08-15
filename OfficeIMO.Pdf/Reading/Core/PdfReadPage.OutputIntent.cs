namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private bool HasOutputIntentCompositionInteraction() {
        if (IsTransparencyGroup(_pageDict)) return true;
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeStreams = new HashSet<PdfStream>();
        var budget = new PageContentBudget(this);
        var type3GlyphBudget = new Type3GlyphBudget(_limits.MaxType3GlyphInvocationsPerPage);
        if (ContentUsesOutputIntentCompositionInteraction(
                GetContentStreamContent(budget),
                resources,
                activeStreams,
                budget,
                type3GlyphBudget,
                0)) return true;

        PdfArray? annotations = ResolveArray(
            _pageDict.Items.TryGetValue("Annots", out PdfObject? annotationsObject) ? annotationsObject : null);
        if (annotations == null) return false;
        EnsureAnnotationBudget(annotations);
        for (int index = 0; index < annotations.Items.Count; index++) {
            PdfDictionary? annotation = ResolveDictionary(annotations.Items[index]);
            if (annotation == null || IsHiddenAnnotation(annotation) || HasNoVisibleAnnotationArea(annotation)) continue;
            if (HasNonDefaultOpacity(annotation, "CA") || HasNonNormalBlendMode(annotation)) return true;
            if (TryGetRenderableAnnotationAppearanceStream(annotation, out PdfStream? appearance, out _) && appearance != null &&
                StreamUsesOutputIntentCompositionInteraction(appearance, resources, activeStreams, budget, type3GlyphBudget, 0)) return true;
        }
        return false;
    }

    private bool ContentUsesOutputIntentCompositionInteraction(
        string content,
        PdfDictionary? resources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        EnsureContentNestingBudget(depth);
        if (Type3TextUsesOutputIntentCompositionInteraction(
                content,
                resources,
                activeStreams,
                budget,
                type3GlyphBudget,
                depth)) return true;
        bool found = false;
        PdfPageOptionalContentVisibility? optionalContentVisibility = GetOptionalContentVisibility(resources);
        var hiddenContentStack = new Stack<bool>();
        var malformedOptionalContentStack = new Stack<bool>();
        PdfContentStreamInterpreter.Interpret(content, _limits.MaxContentOperations, operation => {
            if (found) return;
            if (operation.Name == "BDC") {
                object? tag = operation.Operands.Count > 1
                    ? operation.Operands[operation.Operands.Count - 2]
                    : null;
                object? property = operation.Operands.Count > 0
                    ? operation.Operands[operation.Operands.Count - 1]
                    : null;
                hiddenContentStack.Push(
                    IsHiddenOptionalContent(tag, property, optionalContentVisibility));
                malformedOptionalContentStack.Push(IsMalformedOptionalContent(tag, property));
                return;
            }
            if (operation.Name == "BMC") {
                hiddenContentStack.Push(false);
                malformedOptionalContentStack.Push(false);
                return;
            }
            if (operation.Name == "EMC") {
                if (hiddenContentStack.Count > 0) hiddenContentStack.Pop();
                if (malformedOptionalContentStack.Count > 0) malformedOptionalContentStack.Pop();
                return;
            }
            if (hiddenContentStack.Contains(true)) return;
            if (operation.InlineImage is PdfContentInlineImage inlineImage &&
                HasImageTransparency(inlineImage.Dictionary)) {
                found = true;
                return;
            }
            if (resources == null || operation.Operands.Count == 0) return;
            string? name = operation.Operands[operation.Operands.Count - 1] as string;
            if (name == null) return;
            if (operation.Name == "gs" && malformedOptionalContentStack.Contains(true)) {
                PdfDictionary? states = ResolveDictionary(
                    resources.Items.TryGetValue("ExtGState", out PdfObject? statesObject) ? statesObject : null);
                PdfDictionary? state = states?.Items.TryGetValue(name, out PdfObject? stateObject) == true
                    ? ResolveDictionary(stateObject)
                    : null;
                found = state != null && HasExplicitTransparency(state);
                return;
            }
            if (operation.Name == "Do") {
                PdfDictionary? xObjects = ResolveDictionary(
                    resources.Items.TryGetValue("XObject", out PdfObject? xObjectsObject) ? xObjectsObject : null);
                if (xObjects?.Items.TryGetValue(name, out PdfObject? xObject) == true &&
                    PdfObjectLookup.ResolveChain(_objects, xObject) is PdfStream stream) {
                    found = StreamUsesOutputIntentCompositionInteraction(stream, resources, activeStreams, budget, type3GlyphBudget, depth + 1);
                }
                return;
            }
        },
        inlineImageComponentCount: name => GetDeclaredColorSpaceComponentCount(resources, name),
        maxNestingDepth: _limits.MaxContentNestingDepth,
        maxOperands: _limits.MaxContentOperands,
        inlineImageArrayComponentCount: array => GetDeclaredColorSpaceComponentCount(array));
        return found;

        static bool IsHiddenOptionalContent(
            object? tag,
            object? property,
            PdfPageOptionalContentVisibility? visibility) =>
            tag is string tagName &&
            string.Equals(tagName, "OC", StringComparison.Ordinal) &&
            ((property is string propertyName && visibility?.IsHidden(propertyName) == true) ||
             (property is PdfInlineOptionalContentReferences references && visibility?.IsHidden(references) == true) ||
             (property is PdfContentDictionary dictionary &&
                dictionary.OptionalContentReferences is not null &&
                visibility?.IsHidden(dictionary.OptionalContentReferences) == true));

        static bool IsMalformedOptionalContent(object? tag, object? property) =>
            tag is string tagName &&
            string.Equals(tagName, "OC", StringComparison.Ordinal) &&
            property is not string and
            not PdfInlineOptionalContentReferences and
            not PdfContentDictionary { OptionalContentReferences: not null };
    }

    private bool StreamUsesOutputIntentCompositionInteraction(
        PdfStream stream,
        PdfDictionary? inheritedResources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        string? subtype = (PdfObjectLookup.ResolveChain(
            _objects,
            stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null) as PdfName)?.Name;
        if (subtype == "Image") return HasImageTransparency(stream.Dictionary);
        int? patternType = TryReadInteger(
            stream.Dictionary.Items.TryGetValue("PatternType", out PdfObject? patternTypeObject)
                ? patternTypeObject
                : null);
        if (subtype != "Form" && patternType != 1) return false;
        if (IsTransparencyGroup(stream.Dictionary)) return true;
        if (!activeStreams.Add(stream)) return false;
        try {
            PdfDictionary? resources = ResolveDictionary(
                stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) ??
                inheritedResources;
            return ContentUsesOutputIntentCompositionInteraction(
                PdfEncoding.Latin1GetString(budget.Decode(stream)),
                resources,
                activeStreams,
                budget,
                type3GlyphBudget,
                depth);
        } finally {
            activeStreams.Remove(stream);
        }
    }

    private bool Type3TextUsesOutputIntentCompositionInteraction(
        string content,
        PdfDictionary? resources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        Dictionary<string, PdfFontResource> fonts = resources == null
            ? new Dictionary<string, PdfFontResource>(StringComparer.Ordinal)
            : ResourceResolver.GetFontsForResources(resources, _objects);
        PdfPageInvokedResourceNames invokedResources = GetInvokedResourceNames(content, resources);

        bool found = false;
        PdfPageXObjectInvocationParser.Parse(
            content,
            Matrix2D.Identity,
            GetVisualPageSize().Height,
            GetGraphicsStateResources(resources),
            GetColorSpaceResources(resources, invokedResources.ColorSpaces, budget),
            GetOptionalContentVisibility(resources),
            maxOperations: _limits.MaxContentOperations,
            maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: _limits.MaxContentOperands,
            fonts: fonts,
            fontWidthProviders: resources == null
                ? new Dictionary<string, Func<byte[], double>>(StringComparer.Ordinal)
                : ResourceResolver.GetFontWidthProvidersForResources(resources, _objects),
            type3TextVisitor: invocation => {
                for (int index = 0; index < invocation.Glyphs.Count && !found; index++) {
                    PdfPageType3GlyphInvocation glyph = invocation.Glyphs[index];
                    if (glyph.Font.Type3 is PdfType3FontResource type3 &&
                        type3.TryGetGlyph(glyph.CharacterCode, out PdfStream glyphStream)) {
                        found = Type3GlyphUsesOutputIntentCompositionInteraction(
                            glyphStream,
                            type3.Resources,
                            activeStreams,
                            budget,
                            type3GlyphBudget,
                            depth + 1);
                    }
                }
                return false;
            },
            type3GlyphBudgetConsumer: type3GlyphBudget.Consume,
            patternInvocationVisitor: name => {
                if (found || resources == null) return;
                PdfDictionary? patterns = ResolveDictionary(
                    resources.Items.TryGetValue("Pattern", out PdfObject? patternsObject) ? patternsObject : null);
                if (patterns?.Items.TryGetValue(name, out PdfObject? patternObject) == true &&
                    PdfObjectLookup.ResolveChain(_objects, patternObject) is PdfStream pattern) {
                    found = StreamUsesOutputIntentCompositionInteraction(
                        pattern,
                        resources,
                        activeStreams,
                        budget,
                        type3GlyphBudget,
                        depth + 1);
                }
            },
            graphicsEffectPaintVisitor: (state, channels) => {
                if (HasExplicitTransparency(state, channels)) found = true;
            },
            inlineImageArrayComponentCount: array => GetDeclaredColorSpaceComponentCount(array));
        return found;
    }

    private bool Type3GlyphUsesOutputIntentCompositionInteraction(
        PdfStream stream,
        PdfDictionary resources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        Type3GlyphBudget type3GlyphBudget,
        int depth) {
        if (!activeStreams.Add(stream)) return false;
        try {
            return ContentUsesOutputIntentCompositionInteraction(
                PdfEncoding.Latin1GetString(budget.Decode(stream)),
                resources,
                activeStreams,
                budget,
                type3GlyphBudget,
                depth);
        } finally {
            activeStreams.Remove(stream);
        }
    }

    private static bool HasExplicitTransparency(
        PdfPageGraphicsStateResource state,
        PdfType3PaintChannels channels) {
        if ((channels & PdfType3PaintChannels.Fill) != 0 && state.FillOpacity.HasValue && state.FillOpacity.Value != 1D) return true;
        if ((channels & PdfType3PaintChannels.Stroke) != 0 && state.StrokeOpacity.HasValue && state.StrokeOpacity.Value != 1D) return true;
        if ((state.BlendMode.HasValue && state.BlendMode.Value != OfficeIMO.Drawing.OfficeBlendMode.Normal) ||
            state.HasUnsupportedBlendMode || state.HasUnsupportedSoftMask) return true;
        return state.SoftMaskEnabled == true && state.SoftMask != null;
    }

    private bool HasExplicitTransparency(PdfDictionary dictionary) {
        if (HasNonDefaultOpacity(dictionary, "ca") || HasNonDefaultOpacity(dictionary, "CA") ||
            HasNonNormalBlendMode(dictionary)) return true;
        if (!dictionary.Items.TryGetValue("SMask", out PdfObject? softMask)) return false;
        PdfObject? resolved = PdfObjectLookup.ResolveChain(_objects, softMask);
        return resolved is not PdfNull and not PdfName { Name: "None" };
    }

    private bool HasNonDefaultOpacity(PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) return false;
        return PdfObjectLookup.ResolveChain(_objects, value) is not PdfNumber { Value: 1D };
    }

    private bool HasNonNormalBlendMode(PdfDictionary dictionary) {
        if (!dictionary.Items.ContainsKey("BM")) return false;
        OfficeIMO.Drawing.OfficeBlendMode? blendMode = ReadBlendMode(dictionary);
        return blendMode.HasValue && blendMode.Value != OfficeIMO.Drawing.OfficeBlendMode.Normal;
    }

    private bool HasImageTransparency(PdfDictionary dictionary) {
        if (dictionary.Items.TryGetValue("ImageMask", out PdfObject? imageMask) &&
            PdfObjectLookup.ResolveChain(_objects, imageMask) is PdfBoolean { Value: true }) return true;
        if (dictionary.Items.TryGetValue("SMask", out PdfObject? softMask)) {
            PdfObject? resolved = PdfObjectLookup.ResolveChain(_objects, softMask);
            if (resolved is not PdfNull and not PdfName { Name: "None" }) return true;
        }
        if (!dictionary.Items.TryGetValue("Mask", out PdfObject? mask)) return false;
        return PdfObjectLookup.ResolveChain(_objects, mask) is not PdfNull;
    }

    private bool IsTransparencyGroup(PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("Group", out PdfObject? groupObject) ||
            PdfObjectLookup.ResolveChain(_objects, groupObject) is not PdfDictionary group ||
            !group.Items.TryGetValue("S", out PdfObject? subtype)) return false;
        return PdfObjectLookup.ResolveChain(_objects, subtype) is PdfName { Name: "Transparency" };
    }
}

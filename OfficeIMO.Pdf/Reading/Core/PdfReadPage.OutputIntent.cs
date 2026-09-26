using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal (bool HasDeviceRgb, bool HasDeviceIndependent, bool HasTransparency) GetDefiniteUnlayeredPrintColorUse(
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (HasEffectiveOptionalContentEntry(_pageDict)) return (false, false, false);
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        var budget = new PageContentBudget(this, cancellationToken);
        bool foundRgb = false;
        bool foundIndependent = false;
        bool foundTransparency = false;
        var activeForms = new HashSet<PdfStream>();
        Scan(GetContentStreamContent(budget), resources, false, false, false, false,
            (Fill: false, Stroke: false, Blend: false, SoftMask: false), 0);
        return (foundRgb, foundIndependent, foundTransparency);

        void Scan(string content, PdfDictionary? currentResources, bool initialFillRgb,
            bool initialStrokeRgb, bool initialFillIndependent, bool initialStrokeIndependent,
            (bool Fill, bool Stroke, bool Blend, bool SoftMask) initialTransparency, int depth) {
            EnsureContentNestingBudget(depth);
            PdfDictionary? colorSpaces = ResolveDictionary(
                currentResources?.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject) == true
                    ? colorSpaceObject : null);
            var selectedColorSpaces = new Dictionary<string, (bool UsesDeviceRgb, bool UsesDeviceIndependent)>(StringComparer.Ordinal);
            bool defaultRgbIsOverridden = colorSpaces?.Items.ContainsKey("DefaultRGB") == true;
            bool fillRgb = initialFillRgb;
            bool strokeRgb = initialStrokeRgb;
            bool fillIndependent = initialFillIndependent;
            bool strokeIndependent = initialStrokeIndependent;
            var transparency = initialTransparency;
            int textMode = 0;
            int layeredDepth = 0;
            var markedContent = new Stack<bool>();
            var states = new Stack<(bool FillRgb, bool StrokeRgb,
                bool FillIndependent, bool StrokeIndependent,
                (bool Fill, bool Stroke, bool Blend, bool SoftMask) Transparency, int TextMode)>();
            PdfContentStreamInterpreter.Interpret(content, _limits.MaxContentOperations, operation => {
                cancellationToken.ThrowIfCancellationRequested();
                switch (operation.Name) {
                    case "BDC":
                        bool isLayer = operation.Operands.Count > 1 &&
                            operation.Operands[operation.Operands.Count - 2] is string tag && tag == "OC";
                        markedContent.Push(isLayer);
                        if (isLayer) layeredDepth++;
                        return;
                    case "BMC": markedContent.Push(false); return;
                    case "EMC":
                        if (markedContent.Count > 0 && markedContent.Pop()) layeredDepth--;
                        return;
                }
                if (operation.HasInvalidOperands) return;
                // Optional content suppresses painting, not persistent graphics-state changes.
                if (layeredDepth != 0 && operation.Name is not ("q" or "Q" or "rg" or "RG" or
                    "g" or "G" or "k" or "K" or "cs" or "CS" or "Tr" or "gs")) return;
                if (operation.InlineImage is PdfContentInlineImage inlineImage) {
                    string? inlineColorSpace = (ResolveObject(inlineImage.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? inlineColor)
                        ? inlineColor : null) as PdfName)?.Name;
                    foundRgb |= !defaultRgbIsOverridden && ClassifySelected(inlineColorSpace).UsesDeviceRgb;
                    foundRgb |= fillRgb && inlineImage.Dictionary.Items.TryGetValue("ImageMask", out PdfObject? inlineMask) &&
                        ResolveObject(inlineMask) is PdfBoolean { Value: true };
                    foundIndependent |= fillIndependent && inlineImage.Dictionary.Items.TryGetValue("ImageMask", out PdfObject? independentInlineMask) &&
                        ResolveObject(independentInlineMask) is PdfBoolean { Value: true };
                    foundIndependent |= IsIndependent(inlineColorSpace, colorSpaces) ||
                        IsIndependentColorObject(inlineImage.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? inlineSpace) ? inlineSpace : null);
                    foundTransparency |= transparency.Fill || transparency.Blend ||
                        transparency.SoftMask || HasImageTransparency(inlineImage.Dictionary);
                    return;
                }
                switch (operation.Name) {
                    case "q": states.Push((fillRgb, strokeRgb, fillIndependent, strokeIndependent, transparency, textMode)); break;
                    case "Q":
                        if (states.Count > 0) (fillRgb, strokeRgb, fillIndependent, strokeIndependent, transparency, textMode) = states.Pop();
                        else { fillRgb = strokeRgb = fillIndependent = strokeIndependent = false; transparency = default; textMode = 0; }
                        break;
                    case "rg": fillRgb = !defaultRgbIsOverridden && operation.Operands.Count == 3; fillIndependent = defaultRgbIsOverridden && IsIndependent("DefaultRGB", colorSpaces); break;
                    case "RG": strokeRgb = !defaultRgbIsOverridden && operation.Operands.Count == 3; strokeIndependent = defaultRgbIsOverridden && IsIndependent("DefaultRGB", colorSpaces); break;
                    case "g": case "k": fillRgb = fillIndependent = false; break;
                    case "G": case "K": strokeRgb = strokeIndependent = false; break;
                    case "cs":
                        (fillRgb, fillIndependent) = operation.Operands.Count == 1
                            ? ClassifySelected(operation.Operands[0] as string) : default;
                        break;
                    case "CS":
                        (strokeRgb, strokeIndependent) = operation.Operands.Count == 1
                            ? ClassifySelected(operation.Operands[0] as string) : default;
                        break;
                    case "Tr":
                        if (operation.Operands.Count == 1 && operation.Operands[0] is double mode &&
                            mode >= 0D && mode <= 7D) textMode = (int)mode;
                        break;
                    case "gs":
                        if (operation.Operands.Count > 0 && operation.Operands[operation.Operands.Count - 1] is string stateName) {
                            PdfDictionary? extStates = ResolveDictionary(
                                currentResources?.Items.TryGetValue("ExtGState", out PdfObject? statesObject) == true
                                    ? statesObject : null);
                            PdfDictionary? state = ResolveDictionary(
                                extStates?.Items.TryGetValue(stateName, out PdfObject? stateObject) == true
                                    ? stateObject : null);
                            if (state != null) {
                                if (state.Items.ContainsKey("ca")) transparency.Fill = HasNonDefaultOpacity(state, "ca");
                                if (state.Items.ContainsKey("CA")) transparency.Stroke = HasNonDefaultOpacity(state, "CA");
                                if (state.Items.ContainsKey("BM")) transparency.Blend = HasNonNormalBlendMode(state);
                                if (state.Items.TryGetValue("SMask", out PdfObject? mask)) {
                                    PdfObject? resolvedMask = PdfObjectLookup.ResolveChain(_objects, mask);
                                    transparency.SoftMask = resolvedMask is not PdfNull and not PdfName { Name: "None" };
                                }
                            }
                        }
                        break;
                    case "f": case "F": case "f*": foundRgb |= fillRgb; foundIndependent |= fillIndependent; foundTransparency |= transparency.Fill || transparency.Blend || transparency.SoftMask; break;
                    case "S": case "s": foundRgb |= strokeRgb; foundIndependent |= strokeIndependent; foundTransparency |= transparency.Stroke || transparency.Blend || transparency.SoftMask; break;
                    case "B": case "B*": case "b": case "b*":
                        foundRgb |= fillRgb || strokeRgb;
                        foundIndependent |= fillIndependent || strokeIndependent;
                        foundTransparency |= transparency.Fill || transparency.Stroke || transparency.Blend || transparency.SoftMask;
                        break;
                    case "Tj": case "TJ": case "'": case "\"":
                        if (!GetShownTextBytes(operation).Any(static bytes => bytes.Length > 0)) break;
                        if (textMode is 0 or 2 or 4 or 6) foundRgb |= fillRgb;
                        if (textMode is 1 or 2 or 5 or 6) foundRgb |= strokeRgb;
                        if (textMode is 0 or 2 or 4 or 6) foundIndependent |= fillIndependent;
                        if (textMode is 1 or 2 or 5 or 6) foundIndependent |= strokeIndependent;
                        if (textMode != 3 && textMode != 7) {
                            foundTransparency |= (textMode is 0 or 2 or 4 or 6 && transparency.Fill) ||
                                (textMode is 1 or 2 or 5 or 6 && transparency.Stroke) ||
                                transparency.Blend || transparency.SoftMask;
                        }
                        break;
                    case "Do":
                        if (operation.Operands.Count == 0 || operation.Operands[operation.Operands.Count - 1] is not string resourceName) break;
                        PdfDictionary? xObjects = ResolveDictionary(
                            currentResources?.Items.TryGetValue("XObject", out PdfObject? xObjectsObject) == true
                                ? xObjectsObject : null);
                        if (PdfObjectLookup.ResolveChain(_objects,
                                xObjects?.Items.TryGetValue(resourceName, out PdfObject? xObject) == true ? xObject : null) is not PdfStream stream ||
                            HasEffectiveOptionalContentEntry(stream.Dictionary)) break;
                        string? subtype = (ResolveObject(stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject)
                            ? subtypeObject : null) as PdfName)?.Name;
                        if (subtype == "Image") {
                            string? colorSpace = (ResolveObject(stream.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? imageColor)
                                ? imageColor : null) as PdfName)?.Name;
                            foundRgb |= !defaultRgbIsOverridden && ClassifySelected(colorSpace).UsesDeviceRgb;
                            foundRgb |= fillRgb && stream.Dictionary.Items.TryGetValue("ImageMask", out PdfObject? imageMask) &&
                                ResolveObject(imageMask) is PdfBoolean { Value: true };
                            foundIndependent |= fillIndependent && stream.Dictionary.Items.TryGetValue("ImageMask", out PdfObject? independentImageMask) &&
                                ResolveObject(independentImageMask) is PdfBoolean { Value: true };
                            foundIndependent |= IsIndependent(colorSpace, colorSpaces) ||
                                IsIndependentColorObject(stream.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? imageSpace) ? imageSpace : null);
                            foundTransparency |= HasImageTransparency(stream.Dictionary) || transparency.Fill ||
                                transparency.Blend || transparency.SoftMask;
                        } else if (subtype == "Form" && activeForms.Add(stream)) {
                            try {
                                PdfDictionary? formResources = ResolveDictionary(stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourceObject)
                                    ? formResourceObject : null) ?? currentResources;
                                Scan(PdfEncoding.Latin1GetString(budget.Decode(stream)), formResources,
                                    fillRgb, strokeRgb, fillIndependent, strokeIndependent, transparency, depth + 1);
                            } finally { activeForms.Remove(stream); }
                        }
                        break;
                }
            }, maxNestingDepth: _limits.MaxContentNestingDepth, maxOperands: _limits.MaxContentOperands);

            (bool UsesDeviceRgb, bool UsesDeviceIndependent) ClassifySelected(string? name) {
                if (name == null) return default;
                if (!selectedColorSpaces.TryGetValue(name, out var usage)) {
                    usage = PdfPrintProductionColorInspector.ClassifySelectedColorSpace(name, currentResources,
                        _objects, _limits.MaxObjectNestingDepth, _limits.MaxDecodedStreamBytes);
                    selectedColorSpaces.Add(name, usage);
                }
                return usage;
            }
        }

        bool IsIndependent(string? name, PdfDictionary? colorSpaces) =>
            name != null && colorSpaces?.Items.TryGetValue(name, out PdfObject? selected) == true &&
            IsIndependentColorObject(selected);

        bool IsIndependentColorObject(PdfObject? selected) =>
            PdfPrintProductionColorInspector.UsesDeviceIndependentColorSpace(
                selected, _objects, _limits.MaxObjectNestingDepth, _limits.MaxDecodedStreamBytes);
    }

    private bool GetOutputIntentCompositionInteraction(CancellationToken cancellationToken) =>
        _outputIntentColorTransform != null && _hasOutputIntentCompositionInteraction.GetOrCreate(
            this, static (page, token) => page.ScanOutputIntentCompositionInteraction(token), cancellationToken);

    private void PrepareOutputIntentRendering(CancellationToken cancellationToken) {
        if (_outputIntentColorTransform == null) return;
        _outputIntentColorTransform.Prepare(cancellationToken);
        _ = GetOutputIntentCompositionInteraction(cancellationToken);
    }

    private bool ScanOutputIntentCompositionInteraction(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (IsTransparencyGroup(_pageDict)) return true;
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeStreams = new HashSet<PdfStream>();
        var budget = new PageContentBudget(this, cancellationToken);
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
            cancellationToken.ThrowIfCancellationRequested();
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
            budget.CancellationToken.ThrowIfCancellationRequested();
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
        PdfPageInvokedResourceNames invokedResources = GetInvokedResourceNames(content, resources, budget.CancellationToken.ThrowIfCancellationRequested);

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
            inlineImageArrayComponentCount: array => GetDeclaredColorSpaceComponentCount(array),
            operationCheck: budget.CancellationToken.ThrowIfCancellationRequested);
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

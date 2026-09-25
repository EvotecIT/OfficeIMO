namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private static readonly System.Threading.AsyncLocal<Action?> HiddenOptionalContentInspectionObserver =
        new System.Threading.AsyncLocal<Action?>();
    private static readonly System.Threading.AsyncLocal<Action?> OptionalContentUsageInspectionObserver =
        new System.Threading.AsyncLocal<Action?>();

    internal static Action? HiddenOptionalContentInspectionObserverForTesting {
        get => HiddenOptionalContentInspectionObserver.Value;
        set => HiddenOptionalContentInspectionObserver.Value = value;
    }

    internal bool IsHiddenOptionalContent(PdfDictionary? sourceDictionary) {
        if (sourceDictionary is null ||
            !sourceDictionary.Items.TryGetValue("OC", out PdfObject? optionalContentObject)) {
            return false;
        }

        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        PdfPageOptionalContentVisibility? visibility = GetOptionalContentVisibility(pageResources);
        return visibility?.HasUnsupportedViewUsageApplications == false &&
            visibility.IsHidden(optionalContentObject);
    }

    internal bool HasUnsupportedOptionalContentViewUsageApplications() {
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        return GetOptionalContentVisibility(pageResources)?.HasUnsupportedViewUsageApplications == true;
    }

    internal static Action? OptionalContentUsageInspectionObserverForTesting {
        get => OptionalContentUsageInspectionObserver.Value;
        set => OptionalContentUsageInspectionObserver.Value = value;
    }

    internal bool HasOptionalContentUsage(System.Threading.CancellationToken cancellationToken = default) =>
        HasOptionalContentUsage(hiddenOnly: false, cancellationToken);

    internal bool HasOptionalContentUsage(bool hiddenOnly, System.Threading.CancellationToken cancellationToken,
        bool printableAnnotationsOnly = false, bool unsupportedOnly = false) {
        cancellationToken.ThrowIfCancellationRequested();
        OptionalContentUsageInspectionObserverForTesting?.Invoke();
        if (HasRelevantOptionalContentEntry(_pageDict, hiddenOnly, unsupportedOnly)) return true;

        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeStreams = new HashSet<PdfStream>();
        var budget = new PageContentBudget(this, cancellationToken);
        var type3GlyphBudget = new Type3GlyphBudget(_limits.MaxType3GlyphInvocationsPerPage);
        if (ContentUsesOptionalContent(
                GetContentStreamContent(budget),
                resources,
                activeStreams,
                budget,
                type3GlyphBudget,
                depth: 0, hiddenOnly, unsupportedOnly)) return true;

        PdfArray? annotations = ResolveArray(
            _pageDict.Items.TryGetValue("Annots", out PdfObject? annotationsObject) ? annotationsObject : null);
        if (annotations == null) return false;
        EnsureAnnotationBudget(annotations);
        for (int index = 0; index < annotations.Items.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfDictionary? annotation = ResolveDictionary(annotations.Items[index]);
            if (annotation == null) continue;
            if (printableAnnotationsOnly && !IsPrintableAnnotation(annotation)) continue;
            if (HasRelevantOptionalContentEntry(annotation, hiddenOnly, unsupportedOnly)) return true;
            if (TryGetNormalAppearanceStream(annotation, out PdfStream appearanceStream) &&
                AnnotationAppearanceUsesOptionalContent(
                    appearanceStream,
                    resources,
                    activeStreams,
                    budget,
                    type3GlyphBudget, hiddenOnly, unsupportedOnly)) return true;
        }

        return false;
    }

    private bool IsPrintableAnnotation(PdfDictionary annotation) {
        PdfObject? flagsObject = annotation.Items.TryGetValue("F", out PdfObject? value) ? ResolveObject(value) : null;
        if (flagsObject is not PdfNumber flags) return false;
        int bits = (int)flags.Value;
        return (bits & 4) != 0 && (bits & 3) == 0;
    }

    private bool AnnotationAppearanceUsesOptionalContent(
        PdfStream appearanceStream,
        PdfDictionary? pageResources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        Type3GlyphBudget type3GlyphBudget, bool hiddenOnly, bool unsupportedOnly) {
        if (HasRelevantOptionalContentEntry(appearanceStream.Dictionary, hiddenOnly, unsupportedOnly)) return true;
        if (!activeStreams.Add(appearanceStream)) return false;
        try {
            PdfDictionary? appearanceResources = ResolveDictionary(
                appearanceStream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject)
                    ? resourcesObject
                    : null) ?? pageResources;
            return ContentUsesOptionalContent(
                PdfEncoding.Latin1GetString(budget.Decode(appearanceStream)),
                appearanceResources,
                activeStreams,
                budget,
                type3GlyphBudget,
                depth: 1, hiddenOnly, unsupportedOnly);
        } finally {
            activeStreams.Remove(appearanceStream);
        }
    }

    private bool ContentUsesOptionalContent(
        string content,
        PdfDictionary? resources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        Type3GlyphBudget type3GlyphBudget,
        int depth, bool hiddenOnly, bool unsupportedOnly) {
        EnsureContentNestingBudget(depth);
        bool found = false;
        PdfPageOptionalContentVisibility? visibility = hiddenOnly || unsupportedOnly ? GetOptionalContentVisibility(resources) : null;
        string? fontName = null;
        var fontStack = new Stack<string?>();
        PdfFontResourceSet? fontResources = null;
        Dictionary<PdfDictionary, string> declaredFontNames = GetDeclaredFontNames(resources);
        PdfContentStreamInterpreter.Interpret(content, _limits.MaxContentOperations, operation => {
            budget.CancellationToken.ThrowIfCancellationRequested();
            if (found) return;

            if (operation.Name == "BDC") {
                int tagIndex = operation.Operands.Count - 2;
                if (tagIndex >= 0 && operation.Operands[tagIndex] is string tag &&
                    string.Equals(tag, "OC", StringComparison.Ordinal)) {
                    if (unsupportedOnly ? IsUnsupportedMarkedContent(operation, visibility)
                        : !hiddenOnly || IsHiddenMarkedContent(operation, visibility)) found = true;
                }
                return;
            }

            switch (operation.Name) {
                case "q":
                    fontStack.Push(fontName);
                    return;
                case "Q":
                    fontName = fontStack.Count > 0 ? fontStack.Pop() : null;
                    return;
                case "Tf" when operation.Operands.Count == 2 && operation.Operands[0] is string selectedFont:
                    fontName = selectedFont;
                    return;
                case "Tj": case "TJ": case "'": case "\"":
                    if (fontName is string activeFontName &&
                        (fontResources ??= _fontResourceCache.GetOrCreate(resources, _objects)).Fonts
                            .TryGetValue(activeFontName, out PdfFontResource? font) &&
                        font.Type3 is PdfType3FontResource type3) {
                        foreach (byte[] bytes in GetShownTextBytes(operation)) {
                            for (int index = 0; index < bytes.Length && !found; index++) {
                                type3GlyphBudget.Consume(1);
                                if (type3.TryGetGlyph(bytes[index], out PdfStream glyph)) {
                                    found = Type3GlyphUsesOptionalContent(
                                        glyph,
                                        type3.Resources,
                                        activeStreams,
                                        budget,
                                        type3GlyphBudget,
                                        depth + 1, hiddenOnly, unsupportedOnly);
                                }
                            }
                            if (found) break;
                        }
                    }
                    return;
            }

            if (resources == null || operation.Operands.Count == 0) return;
            string? name = operation.Operands[operation.Operands.Count - 1] as string;
            if (name == null) return;
            if (operation.Name == "Do") {
                PdfDictionary? xObjects = ResolveDictionary(
                    resources.Items.TryGetValue("XObject", out PdfObject? xObjectsObject) ? xObjectsObject : null);
                if (xObjects?.Items.TryGetValue(name, out PdfObject? xObject) == true &&
                    PdfObjectLookup.ResolveChain(_objects, xObject) is PdfStream stream) {
                    found = StreamUsesOptionalContent(stream, resources, activeStreams, budget,
                        type3GlyphBudget, depth + 1, hiddenOnly, unsupportedOnly);
                }
                return;
            }

            if (operation.Name == "gs") {
                PdfDictionary? graphicsStates = ResolveDictionary(
                    resources.Items.TryGetValue("ExtGState", out PdfObject? graphicsStatesObject)
                        ? graphicsStatesObject
                        : null);
                PdfDictionary? graphicsState = ResolveDictionary(
                    graphicsStates?.Items.TryGetValue(name, out PdfObject? graphicsStateObject) == true
                        ? graphicsStateObject
                        : null);
                fontResources ??= _fontResourceCache.GetOrCreate(resources, _objects);
                if (graphicsState != null &&
                    TryReadExtGStateFont(graphicsState, declaredFontNames,
                        fontResources.Decoders, fontResources.WidthProviders, fontResources.Fonts,
                        out string? graphicsStateFont, out _) &&
                    !string.IsNullOrEmpty(graphicsStateFont)) {
                    fontName = graphicsStateFont;
                }
                PdfDictionary? softMask = ResolveDictionary(
                    graphicsState?.Items.TryGetValue("SMask", out PdfObject? softMaskObject) == true
                        ? softMaskObject
                        : null);
                if (PdfObjectLookup.ResolveChain(
                        _objects,
                        softMask?.Items.TryGetValue("G", out PdfObject? groupObject) == true
                            ? groupObject
                            : null) is PdfStream group) {
                    found = StreamUsesOptionalContent(group, resources, activeStreams, budget,
                        type3GlyphBudget, depth + 1, hiddenOnly, unsupportedOnly);
                }
                return;
            }

            if (operation.Name is "scn" or "SCN") {
                PdfDictionary? patterns = ResolveDictionary(
                    resources.Items.TryGetValue("Pattern", out PdfObject? patternsObject) ? patternsObject : null);
                if (patterns?.Items.TryGetValue(name, out PdfObject? patternObject) == true &&
                    PdfObjectLookup.ResolveChain(_objects, patternObject) is PdfStream pattern) {
                    found = StreamUsesOptionalContent(pattern, resources, activeStreams, budget,
                        type3GlyphBudget, depth + 1, hiddenOnly, unsupportedOnly);
                }
            }
        },
        inlineImageComponentCount: name => GetDeclaredColorSpaceComponentCount(resources, name),
        maxNestingDepth: _limits.MaxContentNestingDepth,
        maxOperands: _limits.MaxContentOperands,
        inlineImageArrayComponentCount: array => GetDeclaredColorSpaceComponentCount(array));
        return found;
    }

    private bool Type3GlyphUsesOptionalContent(
        PdfStream glyph,
        PdfDictionary? resources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        Type3GlyphBudget type3GlyphBudget,
        int depth, bool hiddenOnly, bool unsupportedOnly) {
        if (HasRelevantOptionalContentEntry(glyph.Dictionary, hiddenOnly, unsupportedOnly)) return true;
        if (!activeStreams.Add(glyph)) return false;
        try {
            return ContentUsesOptionalContent(
                PdfEncoding.Latin1GetString(budget.Decode(glyph)),
                resources,
                activeStreams,
                budget,
                type3GlyphBudget,
                depth, hiddenOnly, unsupportedOnly);
        } finally {
            activeStreams.Remove(glyph);
        }
    }

    private bool StreamUsesOptionalContent(
        PdfStream stream,
        PdfDictionary? inheritedResources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        Type3GlyphBudget type3GlyphBudget,
        int depth, bool hiddenOnly, bool unsupportedOnly) {
        if (HasRelevantOptionalContentEntry(stream.Dictionary, hiddenOnly, unsupportedOnly)) return true;
        string? subtype = (PdfObjectLookup.ResolveChain(
            _objects,
            stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null) as PdfName)?.Name;
        int? patternType = TryReadInteger(
            stream.Dictionary.Items.TryGetValue("PatternType", out PdfObject? patternTypeObject)
                ? patternTypeObject
                : null);
        if (subtype != "Form" && patternType != 1) return false;
        if (!activeStreams.Add(stream)) return false;
        try {
            PdfDictionary? resources = ResolveDictionary(
                stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) ??
                inheritedResources;
            return ContentUsesOptionalContent(
                PdfEncoding.Latin1GetString(budget.Decode(stream)),
                resources,
                activeStreams,
                budget,
                type3GlyphBudget,
                depth, hiddenOnly, unsupportedOnly);
        } finally {
            activeStreams.Remove(stream);
        }
    }

    private bool HasEffectiveOptionalContentEntry(PdfDictionary dictionary) =>
        dictionary.Items.TryGetValue("OC", out PdfObject? optionalContentObject) &&
        PdfObjectLookup.ResolveChain(_objects, optionalContentObject) is not null and not PdfNull;

    private bool HasRelevantOptionalContentEntry(PdfDictionary dictionary, bool hiddenOnly, bool unsupportedOnly) =>
        unsupportedOnly ? IsUnsupportedOptionalContent(dictionary) :
        hiddenOnly ? IsHiddenOptionalContent(dictionary) : HasEffectiveOptionalContentEntry(dictionary);

    private bool IsUnsupportedOptionalContent(PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("OC", out PdfObject? value)) return false;
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        return GetOptionalContentVisibility(resources)?.IsUnsupported(value) == true;
    }

    internal HashSet<PdfContentOrderKey> GetDefiniteUnlayeredImageContentOrderKeys(System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var keys = new HashSet<PdfContentOrderKey>();
        if (HasEffectiveOptionalContentEntry(_pageDict)) return keys;
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        var budget = new PageContentBudget(this, cancellationToken);
        var activeForms = new HashSet<PdfStream>();
        Scan(GetContentStreamContent(budget), resources, PdfContentOrderKey.Root, 0);
        return keys;

        void Scan(string content, PdfDictionary? currentResources, PdfContentOrderKey prefix, int depth) {
            EnsureContentNestingBudget(depth);
            PdfDictionary? xObjects = ResolveDictionary(currentResources?.Items.TryGetValue("XObject", out PdfObject? xObjectsObject) == true
                ? xObjectsObject : null);
            var markedContent = new Stack<bool>();
            int layeredDepth = 0;
            PdfContentStreamInterpreter.Interpret(content, _limits.MaxContentOperations, operation => {
                cancellationToken.ThrowIfCancellationRequested();
                if (operation.Name == "BDC") {
                    bool isLayer = operation.Operands.Count > 1 &&
                        operation.Operands[operation.Operands.Count - 2] is string tag && tag == "OC";
                    markedContent.Push(isLayer);
                    if (isLayer) layeredDepth++;
                    return;
                }
                if (operation.Name == "BMC") { markedContent.Push(false); return; }
                if (operation.Name == "EMC") {
                    if (markedContent.Count > 0 && markedContent.Pop()) layeredDepth--;
                    return;
                }
                if (layeredDepth != 0 || operation.HasInvalidOperands) return;
                PdfContentOrderKey key = prefix.Append(operation.OperatorOffset);
                if (operation.InlineImage is not null) { keys.Add(key); return; }
                if (operation.Name != "Do" || operation.Operands.Count == 0 ||
                    operation.Operands[operation.Operands.Count - 1] is not string name) return;
                if (PdfObjectLookup.ResolveChain(_objects,
                        xObjects?.Items.TryGetValue(name, out PdfObject? xObject) == true ? xObject : null) is not PdfStream stream ||
                    HasEffectiveOptionalContentEntry(stream.Dictionary)) return;
                string? subtype = (ResolveObject(stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject)
                    ? subtypeObject : null) as PdfName)?.Name;
                if (subtype == "Image") {
                    keys.Add(key);
                } else if (subtype == "Form" && activeForms.Add(stream)) {
                    try {
                        PdfDictionary? formResources = ResolveDictionary(stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourceObject)
                            ? formResourceObject : null) ?? currentResources;
                        Scan(WrapFormContentWithBoundingBoxClip(PdfEncoding.Latin1GetString(budget.Decode(stream)), stream.Dictionary),
                            formResources, key, depth + 1);
                    } finally { activeForms.Remove(stream); }
                }
            }, maxNestingDepth: _limits.MaxContentNestingDepth, maxOperands: _limits.MaxContentOperands);
        }
    }

    private static bool IsHiddenMarkedContent(PdfContentOperation operation, PdfPageOptionalContentVisibility? visibility) {
        object? property = operation.Operands.Count > 0 ? operation.Operands[operation.Operands.Count - 1] : null;
        return (property is string name && visibility?.IsHidden(name) == true) ||
            (property is PdfInlineOptionalContentReferences references && visibility?.IsHidden(references) == true) ||
            (property is PdfContentDictionary dictionary && dictionary.OptionalContentReferences is not null &&
                visibility?.IsHidden(dictionary.OptionalContentReferences) == true);
    }

    private static bool IsUnsupportedMarkedContent(PdfContentOperation operation, PdfPageOptionalContentVisibility? visibility) {
        object? property = operation.Operands.Count > 0 ? operation.Operands[operation.Operands.Count - 1] : null;
        return (property is string name && visibility?.IsUnsupported(name) == true) ||
            (property is PdfInlineOptionalContentReferences references && visibility?.IsUnsupported(references) == true) ||
            (property is PdfContentDictionary dictionary && dictionary.OptionalContentReferences is not null &&
                visibility?.IsUnsupported(dictionary.OptionalContentReferences) == true);
    }

    internal IReadOnlyList<PdfTextSpan> GetHiddenOptionalContentTextSpans(bool includeArtifactText) {
        HiddenOptionalContentInspectionObserver.Value?.Invoke();
        if (HasUnsupportedOptionalContentViewUsageApplications()) {
            return Array.Empty<PdfTextSpan>();
        }

        IReadOnlyList<PdfTextSpan> visible = GetTextSpans(includeArtifactText, default);
        IReadOnlyList<PdfTextSpan> includingHidden = GetTextSpans(
            includeArtifactText,
            default,
            includeHiddenOptionalContent: true);
        var visibleKeys = new HashSet<PdfContentOrderKey>(visible
            .Select(static span => span.ContentOrderKey)
            .OfType<PdfContentOrderKey>());
        return includingHidden
            .Where(span => span.ContentOrderKey is not null && !visibleKeys.Contains(span.ContentOrderKey))
            .ToArray();
    }

    internal IReadOnlyList<PdfTextSpan> GetTextSpansIncludingHiddenOptionalContent() =>
        GetTextSpans(_includeArtifactText, default, includeHiddenOptionalContent: true);
}

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private static readonly System.Threading.AsyncLocal<Action?> HiddenOptionalContentInspectionObserver =
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

    internal bool HasOptionalContentUsage(System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (HasEffectiveOptionalContentEntry(_pageDict)) return true;

        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeStreams = new HashSet<PdfStream>();
        var budget = new PageContentBudget(this, cancellationToken);
        if (ContentUsesOptionalContent(
                GetContentStreamContent(budget),
                resources,
                activeStreams,
                budget,
                depth: 0)) return true;

        PdfArray? annotations = ResolveArray(
            _pageDict.Items.TryGetValue("Annots", out PdfObject? annotationsObject) ? annotationsObject : null);
        if (annotations == null) return false;
        EnsureAnnotationBudget(annotations);
        for (int index = 0; index < annotations.Items.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfDictionary? annotation = ResolveDictionary(annotations.Items[index]);
            if (annotation == null) continue;
            if (HasEffectiveOptionalContentEntry(annotation)) return true;
            if (TryGetNormalAppearanceStream(annotation, out PdfStream appearanceStream) &&
                AnnotationAppearanceUsesOptionalContent(
                    appearanceStream,
                    resources,
                    activeStreams,
                    budget)) return true;
        }

        return false;
    }

    private bool AnnotationAppearanceUsesOptionalContent(
        PdfStream appearanceStream,
        PdfDictionary? pageResources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget) {
        if (HasEffectiveOptionalContentEntry(appearanceStream.Dictionary)) return true;
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
                depth: 1);
        } finally {
            activeStreams.Remove(appearanceStream);
        }
    }

    private bool ContentUsesOptionalContent(
        string content,
        PdfDictionary? resources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        int depth) {
        EnsureContentNestingBudget(depth);
        bool found = false;
        string? fontName = null;
        var fontStack = new Stack<string?>();
        Dictionary<string, PdfFontResource>? fonts = null;
        PdfContentStreamInterpreter.Interpret(content, _limits.MaxContentOperations, operation => {
            budget.CancellationToken.ThrowIfCancellationRequested();
            if (found) return;

            if (operation.Name == "BDC") {
                int tagIndex = operation.Operands.Count - 2;
                if (tagIndex >= 0 && operation.Operands[tagIndex] is string tag &&
                    string.Equals(tag, "OC", StringComparison.Ordinal)) {
                    found = true;
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
                        (fonts ??= ResourceResolver.GetFontsForResources(resources, _objects))
                            .TryGetValue(activeFontName, out PdfFontResource? font) &&
                        font.Type3 is PdfType3FontResource type3) {
                        foreach (byte[] bytes in GetShownTextBytes(operation)) {
                            for (int index = 0; index < bytes.Length && !found; index++) {
                                if (type3.TryGetGlyph(bytes[index], out PdfStream glyph)) {
                                    found = Type3GlyphUsesOptionalContent(
                                        glyph,
                                        type3.Resources,
                                        activeStreams,
                                        budget,
                                        depth + 1);
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
                    found = StreamUsesOptionalContent(stream, resources, activeStreams, budget, depth + 1);
                }
                return;
            }

            if (operation.Name is "scn" or "SCN") {
                PdfDictionary? patterns = ResolveDictionary(
                    resources.Items.TryGetValue("Pattern", out PdfObject? patternsObject) ? patternsObject : null);
                if (patterns?.Items.TryGetValue(name, out PdfObject? patternObject) == true &&
                    PdfObjectLookup.ResolveChain(_objects, patternObject) is PdfStream pattern) {
                    found = StreamUsesOptionalContent(pattern, resources, activeStreams, budget, depth + 1);
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
        int depth) {
        if (HasEffectiveOptionalContentEntry(glyph.Dictionary)) return true;
        if (!activeStreams.Add(glyph)) return false;
        try {
            return ContentUsesOptionalContent(
                PdfEncoding.Latin1GetString(budget.Decode(glyph)),
                resources,
                activeStreams,
                budget,
                depth);
        } finally {
            activeStreams.Remove(glyph);
        }
    }

    private bool StreamUsesOptionalContent(
        PdfStream stream,
        PdfDictionary? inheritedResources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        int depth) {
        if (HasEffectiveOptionalContentEntry(stream.Dictionary)) return true;
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
                depth);
        } finally {
            activeStreams.Remove(stream);
        }
    }

    private bool HasEffectiveOptionalContentEntry(PdfDictionary dictionary) =>
        dictionary.Items.TryGetValue("OC", out PdfObject? optionalContentObject) &&
        PdfObjectLookup.ResolveChain(_objects, optionalContentObject) is not null and not PdfNull;

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

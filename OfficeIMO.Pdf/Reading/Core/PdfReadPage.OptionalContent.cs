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
            if (annotation != null && HasEffectiveOptionalContentEntry(annotation)) return true;
        }

        return false;
    }

    private bool ContentUsesOptionalContent(
        string content,
        PdfDictionary? resources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        int depth) {
        EnsureContentNestingBudget(depth);
        bool found = false;
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

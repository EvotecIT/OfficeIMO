namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private bool HasOutputIntentCompositionInteraction() {
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeStreams = new HashSet<PdfStream>();
        var budget = new PageContentBudget(this);
        if (ContentUsesOutputIntentCompositionInteraction(
                GetContentStreamContent(budget),
                resources,
                activeStreams,
                budget,
                0)) return true;

        PdfArray? annotations = ResolveArray(
            _pageDict.Items.TryGetValue("Annots", out PdfObject? annotationsObject) ? annotationsObject : null);
        if (annotations == null) return false;
        EnsureAnnotationBudget(annotations);
        for (int index = 0; index < annotations.Items.Count; index++) {
            PdfDictionary? annotation = ResolveDictionary(annotations.Items[index]);
            if (annotation == null || IsHiddenAnnotation(annotation) || HasNoVisibleAnnotationArea(annotation)) continue;
            if (HasNonDefaultOpacity(annotation, "CA") || HasNonNormalBlendMode(annotation)) return true;
            if (TryGetNormalAppearanceStream(annotation, out PdfStream? appearance) && appearance != null &&
                StreamUsesOutputIntentCompositionInteraction(appearance, resources, activeStreams, budget, 0)) return true;
        }
        return false;
    }

    private bool ContentUsesOutputIntentCompositionInteraction(
        string content,
        PdfDictionary? resources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        int depth) {
        EnsureContentNestingBudget(depth);
        bool found = false;
        PdfContentStreamInterpreter.Interpret(content, _limits.MaxContentOperations, operation => {
            if (found) return;
            if (operation.InlineImage is PdfContentInlineImage inlineImage &&
                HasImageTransparency(inlineImage.Dictionary)) {
                found = true;
                return;
            }
            if (resources == null || operation.Operands.Count == 0) return;
            string? name = operation.Operands[operation.Operands.Count - 1] as string;
            if (name == null) return;
            if (operation.Name == "gs") {
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
                    found = StreamUsesOutputIntentCompositionInteraction(stream, resources, activeStreams, budget, depth + 1);
                }
                return;
            }
            if (operation.Name is "scn" or "SCN") {
                PdfDictionary? patterns = ResolveDictionary(
                    resources.Items.TryGetValue("Pattern", out PdfObject? patternsObject) ? patternsObject : null);
                if (patterns?.Items.TryGetValue(name, out PdfObject? patternObject) == true &&
                    PdfObjectLookup.ResolveChain(_objects, patternObject) is PdfStream pattern) {
                    found = StreamUsesOutputIntentCompositionInteraction(pattern, resources, activeStreams, budget, depth + 1);
                }
            }
        },
        maxNestingDepth: _limits.MaxContentNestingDepth,
        maxOperands: _limits.MaxContentOperands);
        return found;
    }

    private bool StreamUsesOutputIntentCompositionInteraction(
        PdfStream stream,
        PdfDictionary? inheritedResources,
        HashSet<PdfStream> activeStreams,
        PageContentBudget budget,
        int depth) {
        string? subtype = (PdfObjectLookup.ResolveChain(
            _objects,
            stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null) as PdfName)?.Name;
        if (subtype == "Image") return HasImageTransparency(stream.Dictionary);
        if (subtype != "Form" && stream.Dictionary.Get<PdfNumber>("PatternType")?.Value != 1D) return false;
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
                depth);
        } finally {
            activeStreams.Remove(stream);
        }
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

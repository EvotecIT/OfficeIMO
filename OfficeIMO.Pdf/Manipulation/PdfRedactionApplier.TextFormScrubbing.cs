using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private static TextFormScrubContentResult ScrubFormInvocations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary resources,
        PdfDictionary xObjects,
        string content,
        RedactionTextTarget[] textTargets,
        IReadOnlyDictionary<string, Func<byte[], string>> parentFontDecoders,
        IReadOnlyDictionary<string, Func<byte[], double>> parentFontWidthProviders,
        ISet<string> parentVerticalWritingFonts,
        IReadOnlyList<Matrix2D> parentTransforms,
        PdfTextStateSnapshot initialTextState,
        IReadOnlyDictionary<int, int> referenceCounts,
        HashSet<int> activeForms,
        PdfReadLimits limits,
        HashSet<PdfStream> sourceStreamIdentities,
        ref int nextObjectNumber) {
        bool changed = false;
        string rewrittenContent = content;
        ImageResourceInvocation[] invocations = ExtractImageResourceInvocations(
            content,
            initialTextState,
            ResolveExtGStateFontSelections(objects, resources));
        for (int invocationIndex = invocations.Length - 1; invocationIndex >= 0; invocationIndex--) {
            ImageResourceInvocation invocation = invocations[invocationIndex];
            if (!TryGetFormXObject(objects, xObjects, invocation.Name, out PdfReference reference, out PdfStream formStream) ||
                formStream.DecodingFailed ||
                !activeForms.Add(reference.ObjectNumber)) {
                continue;
            }

            int activeObjectNumber = reference.ObjectNumber;
            try {
                bool repeatedInvocation = invocations.Count(candidate => string.Equals(candidate.Name, invocation.Name, StringComparison.Ordinal)) != 1;
                bool sharedReference = IsCurrentlySharedReference(objects, referenceCounts, reference);
                PdfReference sourceReference = reference;
                if ((repeatedInvocation || sharedReference) &&
                    PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? sourceIndirect)) {
                    reference = CloneIndirectObject(objects, reference, sourceIndirect, ref nextObjectNumber);
                    formStream = (PdfStream)objects[reference.ObjectNumber].Value;
                }

                TextFormScrubContentResult result = ScrubFormContent(
                    objects,
                    resources,
                    reference,
                    formStream,
                    !SameReference(reference, sourceReference),
                    invocation.Transform,
                    textTargets,
                    parentFontDecoders,
                    parentFontWidthProviders,
                    parentVerticalWritingFonts,
                    parentTransforms,
                    invocation.TextState,
                    referenceCounts,
                    activeForms,
                    limits,
                    sourceStreamIdentities,
                    ref nextObjectNumber);
                if (!result.HasChanges) {
                    if (!SameReference(reference, sourceReference)) {
                        objects.Remove(reference.ObjectNumber);
                    }

                    continue;
                }

                if (repeatedInvocation) {
                    string resourceName = CreateUniqueResourceName(xObjects, invocation.Name);
                    xObjects.Items[resourceName] = reference;
                    rewrittenContent = ReplaceInvocationResourceName(rewrittenContent, invocation, resourceName);
                } else if (sharedReference) {
                    xObjects.Items[invocation.Name] = reference;
                }

                changed = true;
            } finally {
                activeForms.Remove(activeObjectNumber);
            }
        }

        return new TextFormScrubContentResult(changed, rewrittenContent);
    }

    private static TextFormScrubContentResult ScrubFormContent(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary inheritedResources,
        PdfReference formReference,
        PdfStream formStream,
        bool isolateResources,
        Matrix2D invocationTransform,
        RedactionTextTarget[] textTargets,
        IReadOnlyDictionary<string, Func<byte[], string>> parentFontDecoders,
        IReadOnlyDictionary<string, Func<byte[], double>> parentFontWidthProviders,
        ISet<string> parentVerticalWritingFonts,
        IReadOnlyList<Matrix2D> parentTransforms,
        PdfTextStateSnapshot inheritedTextState,
        IReadOnlyDictionary<int, int> referenceCounts,
        HashSet<int> activeForms,
        PdfReadLimits limits,
        HashSet<PdfStream> sourceStreamIdentities,
        ref int nextObjectNumber) {
        PdfDictionary formResources = ResolveTextFormResources(objects, inheritedResources, formStream, isolateResources);
        PdfDictionary formXObjects = isolateResources
            ? EnsureResourceXObjects(objects, formResources)
            : ResolveDictionary(objects, formResources.Items.TryGetValue("XObject", out PdfObject? formXObjectObject) ? formXObjectObject : null) ?? new PdfDictionary();
        Dictionary<string, Func<byte[], string>> localDecoders = ResourceResolver.GetFontDecodersForForm(formStream.Dictionary, objects);
        Dictionary<string, Func<byte[], string>> formDecoders = MergeDecoders(parentFontDecoders, localDecoders);
        Dictionary<string, Func<byte[], double>> localWidthProviders = ResourceResolver.GetFontWidthProvidersForResources(formResources, objects);
        Dictionary<string, Func<byte[], double>> formWidthProviders = MergeWidthProviders(
            parentFontWidthProviders,
            localWidthProviders);
        PdfTextStateSnapshot formInitialTextState = PreserveInheritedFormFontState(
            inheritedTextState,
            parentFontDecoders,
            parentFontWidthProviders,
            localDecoders,
            localWidthProviders,
            formDecoders,
            formWidthProviders);
        HashSet<string> formVerticalWritingFonts = GetVerticalWritingFontResources(formResources, objects);
        foreach (string parentFont in parentVerticalWritingFonts) {
            formVerticalWritingFonts.Add(parentFont);
        }
        if (parentVerticalWritingFonts.Contains(inheritedTextState.FontResource) &&
            !string.Equals(formInitialTextState.FontResource, inheritedTextState.FontResource, StringComparison.Ordinal)) {
            formVerticalWritingFonts.Add(formInitialTextState.FontResource);
        }
        Matrix2D[] effectiveTransforms = parentTransforms
            .Select(parent => ApplyFormMatrix(Matrix2D.Multiply(parent, invocationTransform), formStream.Dictionary))
            .ToArray();
        string formContent = PdfEncoding.Latin1GetString(StreamDecoder.DecodeRequired(formStream.Dictionary, formStream.Data, objects, GetMutationDecodeLimit(formStream, limits, sourceStreamIdentities)));
        var formGraphicsState = new TextScrubGraphicsState { TextState = formInitialTextState };
        IReadOnlyDictionary<string, PdfExtGStateFontSelection> formExtGStateFonts = ResolveExtGStateFontSelections(objects, formResources);
        string scrubbed = ScrubTextObjects(formContent, textTargets, formDecoders, formWidthProviders, effectiveTransforms, limits, formGraphicsState, formExtGStateFonts, formVerticalWritingFonts);
        bool changed = !string.Equals(formContent, scrubbed, StringComparison.Ordinal);
        TextFormScrubContentResult nestedResult = ScrubFormInvocations(
            objects,
            formResources,
            formXObjects,
            scrubbed,
            textTargets,
            formDecoders,
            formWidthProviders,
            formVerticalWritingFonts,
            effectiveTransforms,
            formInitialTextState,
            referenceCounts,
            activeForms,
            limits,
            sourceStreamIdentities,
            ref nextObjectNumber);
        string rewrittenContent = nestedResult.Content;
        if (!string.Equals(formContent, rewrittenContent, StringComparison.Ordinal)) {
            objects[formReference.ObjectNumber] = new PdfIndirectObject(
                formReference.ObjectNumber,
                formReference.Generation,
                new PdfStream(CleanStreamDictionary(formStream.Dictionary), PdfEncoding.Latin1GetBytes(rewrittenContent)));
        }

        return new TextFormScrubContentResult(changed || nestedResult.HasChanges, rewrittenContent);
    }

    private static PdfTextStateSnapshot PreserveInheritedFormFontState(
        PdfTextStateSnapshot inheritedTextState,
        IReadOnlyDictionary<string, Func<byte[], string>> parentFontDecoders,
        IReadOnlyDictionary<string, Func<byte[], double>> parentFontWidthProviders,
        Dictionary<string, Func<byte[], string>> localDecoders,
        Dictionary<string, Func<byte[], double>> localWidthProviders,
        Dictionary<string, Func<byte[], string>> formDecoders,
        Dictionary<string, Func<byte[], double>> formWidthProviders) {
        string inheritedFont = inheritedTextState.FontResource;
        if (!localDecoders.ContainsKey(inheritedFont) && !localWidthProviders.ContainsKey(inheritedFont)) {
            return inheritedTextState;
        }

        const string AliasPrefix = "__OfficeIMOInheritedFont";
        string alias = AliasPrefix;
        int suffix = 1;
        while (formDecoders.ContainsKey(alias) || formWidthProviders.ContainsKey(alias)) {
            alias = AliasPrefix + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
            suffix++;
        }
        if (parentFontDecoders.TryGetValue(inheritedFont, out Func<byte[], string>? decoder)) {
            formDecoders[alias] = decoder;
        }
        if (parentFontWidthProviders.TryGetValue(inheritedFont, out Func<byte[], double>? widthProvider)) {
            formWidthProviders[alias] = widthProvider;
        }
        return inheritedTextState.WithFont(alias, inheritedTextState.FontSize);
    }

    private static Dictionary<string, PdfExtGStateFontSelection> ResolveExtGStateFontSelections(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary resources) {
        var result = new Dictionary<string, PdfExtGStateFontSelection>(StringComparer.Ordinal);
        PdfDictionary? states = ResolveDictionary(
            objects,
            resources.Items.TryGetValue("ExtGState", out PdfObject? statesObject) ? statesObject : null);
        if (states == null) return result;
        foreach (KeyValuePair<string, PdfObject> entry in states.Items) {
            PdfDictionary? state = ResolveDictionary(objects, entry.Value);
            if (state == null || !state.Items.TryGetValue("Font", out PdfObject? fontObject)) continue;
            PdfObject? resolvedFont = ResolveRedactionObject(objects, fontObject);
            if (resolvedFont is PdfNull) continue;
            if (resolvedFont is not PdfArray fontArray || fontArray.Items.Count != 2 ||
                ResolveRedactionObject(objects, fontArray.Items[0]) is not PdfName fontName ||
                ResolveRedactionObject(objects, fontArray.Items[1]) is not PdfNumber fontSize ||
                double.IsNaN(fontSize.Value) || double.IsInfinity(fontSize.Value) || fontSize.Value < 0D) {
                result[entry.Key] = PdfExtGStateFontSelection.Invalid;
                continue;
            }
            result[entry.Key] = new PdfExtGStateFontSelection(fontName.Name, fontSize.Value);
        }
        return result;
    }

    private static PdfObject? ResolveRedactionObject(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        if (value is not PdfReference reference) return value;
        return PdfObjectLookup.TryResolveReferenceChain(objects, reference, out PdfObject? resolved) ? resolved : null;
    }

    private static PdfDictionary ResolveTextFormResources(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary inheritedResources,
        PdfStream formStream,
        bool isolateResources) {
        if (!isolateResources) {
            return ResolveDictionary(
                objects,
                formStream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) ?? inheritedResources;
        }

        if (formStream.Dictionary.Items.ContainsKey("Resources")) {
            return EnsureFormResources(objects, formStream);
        }

        PdfDictionary resources = CloneDictionary(inheritedResources);
        formStream.Dictionary.Items["Resources"] = resources;
        return resources;
    }

    private readonly struct TextFormScrubContentResult {
        internal TextFormScrubContentResult(bool hasChanges, string content) {
            HasChanges = hasChanges;
            Content = content;
        }

        internal bool HasChanges { get; }

        internal string Content { get; }
    }
}

internal readonly struct PdfExtGStateFontSelection {
    internal PdfExtGStateFontSelection(string fontResource, double fontSize) {
        FontResource = fontResource;
        FontSize = fontSize;
        IsValid = true;
    }

    internal string FontResource { get; }

    internal double FontSize { get; }

    internal bool IsValid { get; }

    internal static PdfExtGStateFontSelection Invalid => default;
}

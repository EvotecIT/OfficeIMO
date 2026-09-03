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
        IReadOnlyList<Matrix2D> parentTransforms,
        PdfTextStateSnapshot initialTextState,
        IReadOnlyDictionary<int, int> referenceCounts,
        HashSet<int> activeForms,
        PdfReadLimits limits,
        HashSet<PdfStream> sourceStreamIdentities,
        ref int nextObjectNumber) {
        bool changed = false;
        string rewrittenContent = content;
        ImageResourceInvocation[] invocations = ExtractImageResourceInvocations(content, initialTextState);
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
        Dictionary<string, Func<byte[], string>> formDecoders = MergeDecoders(parentFontDecoders, ResourceResolver.GetFontDecodersForForm(formStream.Dictionary, objects));
        Dictionary<string, Func<byte[], double>> formWidthProviders = MergeWidthProviders(
            parentFontWidthProviders,
            ResourceResolver.GetFontWidthProvidersForResources(formResources, objects));
        Matrix2D[] effectiveTransforms = parentTransforms
            .Select(parent => ApplyFormMatrix(Matrix2D.Multiply(parent, invocationTransform), formStream.Dictionary))
            .ToArray();
        string formContent = PdfEncoding.Latin1GetString(StreamDecoder.DecodeRequired(formStream.Dictionary, formStream.Data, objects, GetMutationDecodeLimit(formStream, limits, sourceStreamIdentities)));
        var formGraphicsState = new TextScrubGraphicsState { TextState = inheritedTextState };
        string scrubbed = ScrubTextObjects(formContent, textTargets, formDecoders, formWidthProviders, effectiveTransforms, limits, formGraphicsState);
        bool changed = !string.Equals(formContent, scrubbed, StringComparison.Ordinal);
        TextFormScrubContentResult nestedResult = ScrubFormInvocations(
            objects,
            formResources,
            formXObjects,
            scrubbed,
            textTargets,
            formDecoders,
            formWidthProviders,
            effectiveTransforms,
            inheritedTextState,
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

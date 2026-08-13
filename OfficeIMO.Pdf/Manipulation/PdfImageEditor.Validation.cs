using System.Security.Cryptography;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfImageEditor {
    private static PdfImagePlacement[] BindSourceIdentity(
        PdfImagePlacement[] placements,
        byte[] pdf) {
        string identity = ComputeSourceIdentity(pdf);
        for (int i = 0; i < placements.Length; i++) placements[i].SourceDocumentIdentity = identity;
        return placements;
    }

    private static string ComputeSourceIdentity(byte[] pdf) {
#if NET8_0_OR_GREATER
        return Convert.ToBase64String(SHA256.HashData(pdf));
#else
        using SHA256 sha256 = SHA256.Create();
        return Convert.ToBase64String(sha256.ComputeHash(pdf));
#endif
    }

    private static void EnsureReplaceableSourceImage(
        byte[] pdf,
        PdfImagePlacement placement,
        PdfReadOptions? readOptions) {
        if (placement.ObjectNumber <= 0) {
            throw new NotSupportedException("Replacing a direct image XObject is not supported because its source semantics cannot be verified safely.");
        }
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(pdf, readOptions).Map;
        if (!objects.TryGetValue(placement.ObjectNumber, out PdfIndirectObject? indirect) ||
            indirect.Value is not PdfStream imageStream ||
            !string.Equals(imageStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Image", StringComparison.Ordinal)) {
            throw new NotSupportedException("The selected source image XObject could not be verified safely.");
        }
        if (imageStream.Dictionary.Get<PdfBoolean>("ImageMask")?.Value == true) {
            throw new NotSupportedException("Replacing an ImageMask placement is not supported because its paint color is owned by page graphics state.");
        }
    }

    private static void EnsureSafeDestructiveContext(
        byte[] pdf,
        PdfImagePlacement placement,
        bool requirePortableSourceSemantics,
        PdfReadOptions? readOptions) {
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(pdf, readOptions).Map;
        PdfReadLimits limits = readOptions?.Limits ?? new PdfReadLimits();
        int maximumDecodedStreamBytes = limits.MaxDecodedStreamBytes;

        List<(int ObjectNumber, PdfStream Stream, string Content)> decodedStreams = DecodeReachableContentStreams(
            objects,
            limits,
            maximumDecodedStreamBytes,
            out Dictionary<int, HashSet<PdfDictionary>> effectiveResourceOwners);
        HashSet<int> containingForms = CollectContainingForms(objects, decodedStreams, effectiveResourceOwners, placement, limits);
        if (HasStructureTreeAssociation(objects, placement, containingForms)) {
            throw new NotSupportedException("Editing an image associated with the structure tree is not supported because its tagged-content relationship cannot be preserved safely.");
        }
        if (requirePortableSourceSemantics && HasOptionalContentMembership(objects, placement, containingForms)) {
            throw new NotSupportedException("Replacing or moving an image XObject with optional-content membership is not supported because restamping cannot preserve that membership.");
        }
        if (requirePortableSourceSemantics && (placement.HasAuthoredRenderingIntent || HasAuthoredRenderingIntent(objects, placement))) {
            throw new NotSupportedException("Replacing or moving an image XObject with an authored rendering intent is not supported because restamping cannot preserve its color-conversion semantics.");
        }

        if (InvokesSelectedTargetInsideMarkedContentAcrossPageStreams(objects, decodedStreams, placement, containingForms, limits)) {
            throw new NotSupportedException("Editing an image inside tagged, artifact, or optional marked content is not supported because its structural context cannot be preserved safely.");
        }
        for (int i = 0; i < decodedStreams.Count; i++) {
            (int objectNumber, PdfStream stream, string content) = decodedStreams[i];
            if (string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal) &&
                InvokesSelectedTargetInsideMarkedContent(
                    content,
                    name => ResourceMapsToSelectedTarget(objects, GetEffectiveResourceOwners(effectiveResourceOwners, objectNumber, stream.Dictionary), name, placement, containingForms),
                    limits)) {
                throw new NotSupportedException("Editing an image inside tagged, artifact, or optional marked content is not supported because its structural context cannot be preserved safely.");
            }
        }

        EnsureFormInvocationIsIsolated(objects, decodedStreams, effectiveResourceOwners, placement, containingForms, limits);
    }

    private static List<(int ObjectNumber, PdfStream Stream, string Content)> DecodeReachableContentStreams(
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadLimits limits,
        int maximumDecodedStreamBytes,
        out Dictionary<int, HashSet<PdfDictionary>> effectiveResourceOwners) {
        var decodedStreams = new List<(int ObjectNumber, PdfStream Stream, string Content)>();
        var decodedObjectNumbers = new HashSet<int>();
        var pending = new Queue<int>();
        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is not PdfDictionary page ||
                !string.Equals(page.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal) ||
                !page.Items.TryGetValue("Contents", out PdfObject? contents)) continue;
            foreach (PdfReference reference in EnumeratePageContentReferences(contents, objects)) {
                if (!decodedObjectNumbers.Contains(reference.ObjectNumber)) pending.Enqueue(reference.ObjectNumber);
            }
        }

        effectiveResourceOwners = new Dictionary<int, HashSet<PdfDictionary>>();
        while (pending.Count > 0) {
            int objectNumber = pending.Dequeue();
            if (!decodedObjectNumbers.Add(objectNumber) ||
                !objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfStream stream ||
                stream.DecodingFailed) continue;
            try {
                string content = PdfEncoding.Latin1GetString(StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, maximumDecodedStreamBytes));
                decodedStreams.Add((objectNumber, stream, content));
            } catch (NotSupportedException) {
                throw new NotSupportedException("Editing this image is not supported because a reachable content stream cannot be decoded safely.");
            }

            effectiveResourceOwners = BuildEffectiveResourceOwners(objects, decodedStreams, limits);
            foreach (int reachableObjectNumber in effectiveResourceOwners.Keys) {
                if (!decodedObjectNumbers.Contains(reachableObjectNumber) &&
                    objects.TryGetValue(reachableObjectNumber, out PdfIndirectObject? reachable) &&
                    reachable.Value is PdfStream reachableStream &&
                    string.Equals(reachableStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) {
                    pending.Enqueue(reachableObjectNumber);
                }
            }
        }

        effectiveResourceOwners = BuildEffectiveResourceOwners(objects, decodedStreams, limits);
        return decodedStreams;
    }

    private static bool HasAuthoredRenderingIntent(
        Dictionary<int, PdfIndirectObject> objects,
        PdfImagePlacement placement) =>
        placement.ObjectNumber > 0 &&
        objects.TryGetValue(placement.ObjectNumber, out PdfIndirectObject? indirect) &&
        indirect.Value is PdfStream image &&
        image.Dictionary.Items.TryGetValue("Intent", out PdfObject? intent) &&
        PdfObjectLookup.Resolve(objects, intent) is not PdfNull;

    private static bool InvokesSelectedTargetInsideMarkedContent(
        string content,
        Func<string, bool> resolvesToSelectedTarget,
        PdfReadLimits limits) {
        int markedContentDepth = 0;
        return InvokesSelectedTargetInsideMarkedContent(content, resolvesToSelectedTarget, limits, ref markedContentDepth);
    }

    private static bool InvokesSelectedTargetInsideMarkedContent(
        string content,
        Func<string, bool> resolvesToSelectedTarget,
        PdfReadLimits limits,
        ref int markedContentDepth) {
        bool found = false;
        int depth = markedContentDepth;
        PdfContentStreamInterpreter.Interpret(
            content,
            limits.MaxContentOperations,
            operation => {
                if (string.Equals(operation.Name, "BDC", StringComparison.Ordinal) ||
                    string.Equals(operation.Name, "BMC", StringComparison.Ordinal)) {
                    depth++;
                } else if (string.Equals(operation.Name, "EMC", StringComparison.Ordinal)) {
                    if (depth > 0) depth--;
                } else if (depth > 0 &&
                           string.Equals(operation.Name, "Do", StringComparison.Ordinal) &&
                           operation.Operands.Count > 0 &&
                           operation.Operands[operation.Operands.Count - 1] is string name &&
                           resolvesToSelectedTarget(name)) {
                    found = true;
                }
            },
            maxNestingDepth: limits.MaxContentNestingDepth,
            maxOperands: limits.MaxContentOperands);
        markedContentDepth = depth;
        return found;
    }

    private static bool InvokesSelectedTargetInsideMarkedContentAcrossPageStreams(
        Dictionary<int, PdfIndirectObject> objects,
        List<(int ObjectNumber, PdfStream Stream, string Content)> decodedStreams,
        PdfImagePlacement placement,
        HashSet<int> containingForms,
        PdfReadLimits limits) {
        var contentByObjectNumber = decodedStreams.ToDictionary(static item => item.ObjectNumber, static item => item.Content);
        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is not PdfDictionary page ||
                !string.Equals(page.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal) ||
                !page.Items.TryGetValue("Contents", out PdfObject? contents)) continue;
            int markedContentDepth = 0;
            foreach (PdfReference reference in EnumeratePageContentReferences(contents, objects)) {
                if (contentByObjectNumber.TryGetValue(reference.ObjectNumber, out string? content) &&
                    InvokesSelectedTargetInsideMarkedContent(
                        content,
                        name => ResourceMapsToSelectedTarget(objects, page, name, placement, containingForms),
                        limits,
                        ref markedContentDepth)) return true;
            }
        }
        return false;
    }

    private static IEnumerable<PdfReference> EnumeratePageContentReferences(
        PdfObject contents,
        Dictionary<int, PdfIndirectObject> objects) {
        if (contents is PdfReference reference) {
            if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) && indirect.Value is PdfArray array) {
                for (int i = 0; i < array.Items.Count; i++) {
                    foreach (PdfReference item in EnumeratePageContentReferences(array.Items[i], objects)) yield return item;
                }
            } else {
                yield return reference;
            }
        } else if (contents is PdfArray directArray) {
            for (int i = 0; i < directArray.Items.Count; i++) {
                foreach (PdfReference item in EnumeratePageContentReferences(directArray.Items[i], objects)) yield return item;
            }
        }
    }

    private static void EnsureFormInvocationIsIsolated(
        Dictionary<int, PdfIndirectObject> objects,
        List<(int ObjectNumber, PdfStream Stream, string Content)> decodedStreams,
        Dictionary<int, HashSet<PdfDictionary>> effectiveResourceOwners,
        PdfImagePlacement placement,
        HashSet<int> containingForms,
        PdfReadLimits limits) {
        if (containingForms.Count == 0) return;

        var streamsByForm = new Dictionary<int, HashSet<int>>();
        for (int i = 0; i < decodedStreams.Count; i++) {
            foreach (string name in ReadInvokedResourceNames(decodedStreams[i].Content, limits)) {
                foreach (int formObjectNumber in containingForms) {
                    if (!ResourceMapsToObject(objects, GetEffectiveResourceOwners(effectiveResourceOwners, decodedStreams[i].ObjectNumber, decodedStreams[i].Stream.Dictionary), name, formObjectNumber)) continue;
                    if (!streamsByForm.TryGetValue(formObjectNumber, out HashSet<int>? streams)) {
                        streams = new HashSet<int>();
                        streamsByForm[formObjectNumber] = streams;
                    }
                    streams.Add(decodedStreams[i].ObjectNumber);
                }
            }
        }
        if (streamsByForm.Values.Any(static streams => streams.Count > 1)) {
            throw new NotSupportedException("Editing an image inside a Form XObject reused from multiple content streams is not supported because the selected invocation cannot be isolated safely.");
        }
        foreach (HashSet<int> streams in streamsByForm.Values) {
            if (streams.Any(stream => CountContentStreamOwners(objects, stream) > 1)) {
                throw new NotSupportedException("Editing an image inside a Form XObject invoked by shared page content is not supported because the selected invocation cannot be isolated safely.");
            }
        }
    }

    private static HashSet<int> CollectContainingForms(
        Dictionary<int, PdfIndirectObject> objects,
        List<(int ObjectNumber, PdfStream Stream, string Content)> decodedStreams,
        Dictionary<int, HashSet<PdfDictionary>> effectiveResourceOwners,
        PdfImagePlacement placement,
        PdfReadLimits limits) {
        var containingForms = new HashSet<int>();
        for (int i = 0; i < decodedStreams.Count; i++) {
            (int objectNumber, PdfStream stream, string content) = decodedStreams[i];
            if (string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal) &&
                ResourceMapsToSelectedImage(objects, GetEffectiveResourceOwners(effectiveResourceOwners, objectNumber, stream.Dictionary), placement.ResourceName, placement) &&
                InvokesResource(content, placement.ResourceName, limits)) containingForms.Add(objectNumber);
        }

        bool changed;
        do {
            changed = false;
            for (int i = 0; i < decodedStreams.Count; i++) {
                (int objectNumber, PdfStream stream, string content) = decodedStreams[i];
                if (containingForms.Contains(objectNumber) ||
                    !string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) continue;
                foreach (string name in ReadInvokedResourceNames(content, limits)) {
                    if (!ResourceMapsToAnyObject(objects, GetEffectiveResourceOwners(effectiveResourceOwners, objectNumber, stream.Dictionary), name, containingForms)) continue;
                    changed = containingForms.Add(objectNumber) || changed;
                    break;
                }
            }
        } while (changed);
        return containingForms;
    }

    private static Dictionary<int, HashSet<PdfDictionary>> BuildEffectiveResourceOwners(
        Dictionary<int, PdfIndirectObject> objects,
        List<(int ObjectNumber, PdfStream Stream, string Content)> decodedStreams,
        PdfReadLimits limits) {
        var result = new Dictionary<int, HashSet<PdfDictionary>>();
        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is not PdfDictionary page ||
                !string.Equals(page.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal) ||
                !page.Items.TryGetValue("Contents", out PdfObject? contents)) continue;
            foreach (PdfReference reference in EnumeratePageContentReferences(contents, objects)) AddEffectiveResourceOwner(result, reference.ObjectNumber, page);
        }
        for (int index = 0; index < decodedStreams.Count; index++) {
            if (decodedStreams[index].Stream.Dictionary.Items.ContainsKey("Resources")) {
                AddEffectiveResourceOwner(result, decodedStreams[index].ObjectNumber, decodedStreams[index].Stream.Dictionary);
            }
        }

        bool changed;
        do {
            changed = false;
            for (int index = 0; index < decodedStreams.Count; index++) {
                (int callerObjectNumber, PdfStream caller, string content) = decodedStreams[index];
                if (!result.TryGetValue(callerObjectNumber, out HashSet<PdfDictionary>? owners)) continue;
                PdfDictionary[] ownerSnapshot = owners.ToArray();
                foreach (string name in ReadInvokedResourceNames(content, limits)) {
                    for (int ownerIndex = 0; ownerIndex < ownerSnapshot.Length; ownerIndex++) {
                        if (!TryGetXObject(objects, ownerSnapshot[ownerIndex], name, out PdfObject? target) ||
                            target is not PdfReference targetReference ||
                            !objects.TryGetValue(targetReference.ObjectNumber, out PdfIndirectObject? targetIndirect) ||
                            targetIndirect.Value is not PdfStream targetStream ||
                            !string.Equals(targetStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) continue;
                        PdfDictionary effectiveOwner = targetStream.Dictionary.Items.ContainsKey("Resources")
                            ? targetStream.Dictionary
                            : ownerSnapshot[ownerIndex];
                        changed = AddEffectiveResourceOwner(result, targetReference.ObjectNumber, effectiveOwner) || changed;
                    }
                }
            }
        } while (changed);
        return result;
    }

    private static bool AddEffectiveResourceOwner(Dictionary<int, HashSet<PdfDictionary>> owners, int objectNumber, PdfDictionary owner) {
        if (!owners.TryGetValue(objectNumber, out HashSet<PdfDictionary>? values)) {
            values = new HashSet<PdfDictionary>();
            owners.Add(objectNumber, values);
        }
        return values.Add(owner);
    }

    private static IEnumerable<PdfDictionary> GetEffectiveResourceOwners(
        Dictionary<int, HashSet<PdfDictionary>> owners,
        int objectNumber,
        PdfDictionary fallback) => owners.TryGetValue(objectNumber, out HashSet<PdfDictionary>? values) && values.Count > 0
            ? values
            : new[] { fallback };

    private static int CountContentStreamOwners(Dictionary<int, PdfIndirectObject> objects, int streamObjectNumber) {
        int owners = 0;
        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is not PdfDictionary dictionary ||
                !dictionary.Items.TryGetValue("Contents", out PdfObject? contents)) continue;
            if (ContainsContentStreamReference(objects, contents, streamObjectNumber)) owners++;
        }
        return owners;
    }

    private static bool ContainsContentStreamReference(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject contents,
        int streamObjectNumber) {
        if (contents is PdfReference reference) {
            if (reference.ObjectNumber == streamObjectNumber) return true;
            contents = PdfObjectLookup.Resolve(objects, contents) ?? contents;
        }
        if (contents is not PdfArray array) return false;
        for (int i = 0; i < array.Items.Count; i++) {
            if (array.Items[i] is PdfReference item && item.ObjectNumber == streamObjectNumber) return true;
        }
        return false;
    }

    private static bool ResourceMapsToObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary owner,
        string resourceName,
        int objectNumber) {
        return TryGetXObject(objects, owner, resourceName, out PdfObject? value) &&
               value is PdfReference reference && reference.ObjectNumber == objectNumber;
    }

    private static bool ResourceMapsToObject(
        Dictionary<int, PdfIndirectObject> objects,
        IEnumerable<PdfDictionary> owners,
        string resourceName,
        int objectNumber) => owners.Any(owner => ResourceMapsToObject(objects, owner, resourceName, objectNumber));

    private static bool ResourceMapsToSelectedImage(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary owner,
        string resourceName,
        PdfImagePlacement placement) {
        if (!TryGetXObject(objects, owner, resourceName, out PdfObject? value)) return false;
        if (placement.ObjectNumber > 0) {
            return value is PdfReference reference && reference.ObjectNumber == placement.ObjectNumber;
        }
        return value is PdfStream directStream &&
               PdfDirectStreamIdentity.Compute(directStream) == placement.DirectStreamIdentity;
    }

    private static bool ResourceMapsToSelectedImage(
        Dictionary<int, PdfIndirectObject> objects,
        IEnumerable<PdfDictionary> owners,
        string resourceName,
        PdfImagePlacement placement) => owners.Any(owner => ResourceMapsToSelectedImage(objects, owner, resourceName, placement));

    private static bool ResourceMapsToSelectedTarget(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary owner,
        string resourceName,
        PdfImagePlacement placement,
        HashSet<int> containingForms) {
        if (!TryGetXObject(objects, owner, resourceName, out PdfObject? value)) return false;
        if (value is PdfReference reference) {
            return containingForms.Contains(reference.ObjectNumber) ||
                   placement.ObjectNumber > 0 && reference.ObjectNumber == placement.ObjectNumber;
        }
        return placement.ObjectNumber == 0 &&
               value is PdfStream directStream &&
               PdfDirectStreamIdentity.Compute(directStream) == placement.DirectStreamIdentity;
    }

    private static bool ResourceMapsToSelectedTarget(
        Dictionary<int, PdfIndirectObject> objects,
        IEnumerable<PdfDictionary> owners,
        string resourceName,
        PdfImagePlacement placement,
        HashSet<int> containingForms) => owners.Any(owner => ResourceMapsToSelectedTarget(objects, owner, resourceName, placement, containingForms));

    private static bool ContentStreamMapsToObject(
        Dictionary<int, PdfIndirectObject> objects,
        int streamObjectNumber,
        PdfDictionary streamDictionary,
        string resourceName,
        int objectNumber) {
        if (ResourceMapsToObject(objects, streamDictionary, resourceName, objectNumber)) return true;
        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is PdfDictionary page &&
                string.Equals(page.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal) &&
                page.Items.TryGetValue("Contents", out PdfObject? contents) &&
                ContainsContentStreamReference(objects, contents, streamObjectNumber) &&
                ResourceMapsToObject(objects, page, resourceName, objectNumber)) return true;
        }
        return false;
    }

    private static bool ResourceMapsToAnyObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary owner,
        string resourceName,
        HashSet<int> objectNumbers) {
        return TryGetXObject(objects, owner, resourceName, out PdfObject? value) &&
               value is PdfReference reference && objectNumbers.Contains(reference.ObjectNumber);
    }

    private static bool ResourceMapsToAnyObject(
        Dictionary<int, PdfIndirectObject> objects,
        IEnumerable<PdfDictionary> owners,
        string resourceName,
        HashSet<int> objectNumbers) => owners.Any(owner => ResourceMapsToAnyObject(objects, owner, resourceName, objectNumbers));

    private static bool TryGetXObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary owner,
        string resourceName,
        out PdfObject? value) {
        value = null;
        var visited = new HashSet<int>();
        PdfDictionary? current = owner;
        while (current != null) {
            if (current.Items.TryGetValue("Resources", out PdfObject? resourcesObject) &&
                PdfObjectLookup.Resolve(objects, resourcesObject) is PdfDictionary resources &&
                resources.Items.TryGetValue("XObject", out PdfObject? xObjectsObject) &&
                PdfObjectLookup.Resolve(objects, xObjectsObject) is PdfDictionary xObjects &&
                xObjects.Items.TryGetValue(resourceName, out value)) return true;
            if (!current.Items.TryGetValue("Parent", out PdfObject? parentObject)) break;
            if (parentObject is PdfReference reference && !visited.Add(reference.ObjectNumber)) break;
            current = PdfObjectLookup.Resolve(objects, parentObject) as PdfDictionary;
        }
        return false;
    }

    private static bool HasOptionalContentMembership(
        Dictionary<int, PdfIndirectObject> objects,
        PdfImagePlacement placement,
        HashSet<int> containingForms) {
        if (placement.ObjectNumber > 0 &&
            objects.TryGetValue(placement.ObjectNumber, out PdfIndirectObject? sourceIndirect) &&
            sourceIndirect.Value is PdfStream sourceStream &&
            sourceStream.Dictionary.Items.ContainsKey("OC")) return true;
        foreach (int formObjectNumber in containingForms) {
            if (objects.TryGetValue(formObjectNumber, out PdfIndirectObject? formIndirect) &&
                formIndirect.Value is PdfStream formStream &&
                formStream.Dictionary.Items.ContainsKey("OC")) return true;
        }
        if (placement.ObjectNumber != 0) return false;
        foreach (PdfIndirectObject indirect in objects.Values) {
            PdfDictionary? owner = indirect.Value is PdfDictionary dictionary
                ? dictionary
                : indirect.Value is PdfStream stream ? stream.Dictionary : null;
            if (owner != null &&
                TryGetXObject(objects, owner, placement.ResourceName, out PdfObject? value) &&
                value is PdfStream directStream &&
                PdfDirectStreamIdentity.Compute(directStream) == placement.DirectStreamIdentity &&
                directStream.Dictionary.Items.ContainsKey("OC")) return true;
        }
        return false;
    }

    private static bool HasStructureTreeAssociation(
        Dictionary<int, PdfIndirectObject> objects,
        PdfImagePlacement placement,
        HashSet<int> containingForms) {
        if (placement.ObjectNumber > 0 &&
            objects.TryGetValue(placement.ObjectNumber, out PdfIndirectObject? sourceIndirect) &&
            sourceIndirect.Value is PdfStream sourceStream &&
            sourceStream.Dictionary.Items.ContainsKey("StructParent")) return true;
        foreach (int formObjectNumber in containingForms) {
            if (objects.TryGetValue(formObjectNumber, out PdfIndirectObject? formIndirect) &&
                formIndirect.Value is PdfStream formStream &&
                formStream.Dictionary.Items.ContainsKey("StructParent")) return true;
        }
        if (placement.ObjectNumber != 0) return false;
        foreach (PdfIndirectObject indirect in objects.Values) {
            PdfDictionary? owner = indirect.Value is PdfDictionary dictionary
                ? dictionary
                : indirect.Value is PdfStream stream ? stream.Dictionary : null;
            if (owner != null &&
                TryGetXObject(objects, owner, placement.ResourceName, out PdfObject? value) &&
                value is PdfStream directStream &&
                PdfDirectStreamIdentity.Compute(directStream) == placement.DirectStreamIdentity &&
                directStream.Dictionary.Items.ContainsKey("StructParent")) return true;
        }
        return false;
    }

    private static bool InvokesResource(string content, string resourceName, PdfReadLimits limits) =>
        ReadInvokedResourceNames(content, limits).Any(name => string.Equals(name, resourceName, StringComparison.Ordinal));

    private static List<string> ReadInvokedResourceNames(string content, PdfReadLimits limits) {
        var names = new List<string>();
        PdfContentStreamInterpreter.Interpret(
            content,
            limits.MaxContentOperations,
            operation => {
                if (string.Equals(operation.Name, "Do", StringComparison.Ordinal) &&
                    operation.Operands.Count > 0 &&
                    operation.Operands[operation.Operands.Count - 1] is string name) names.Add(name);
            },
            maxNestingDepth: limits.MaxContentNestingDepth,
            maxOperands: limits.MaxContentOperands);
        return names;
    }
}

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
        if (requirePortableSourceSemantics && placement.ObjectNumber > 0 &&
            objects.TryGetValue(placement.ObjectNumber, out PdfIndirectObject? sourceIndirect) &&
            sourceIndirect.Value is PdfStream sourceStream &&
            sourceStream.Dictionary.Items.ContainsKey("OC")) {
            throw new NotSupportedException("Replacing or moving an image XObject with optional-content membership is not supported because restamping cannot preserve that membership.");
        }

        var decodedStreams = new List<(int ObjectNumber, PdfStream Stream, string Content)>();
        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is not PdfStream stream ||
                string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Image", StringComparison.Ordinal) ||
                stream.DecodingFailed) continue;
            string content;
            try {
                content = PdfEncoding.Latin1GetString(StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects));
            } catch (NotSupportedException) {
                throw new NotSupportedException("Editing this image is not supported because a potentially related content stream cannot be decoded safely.");
            }
            decodedStreams.Add((indirect.ObjectNumber, stream, content));
        }

        HashSet<string> structurallyUnsafeNames = ReadStructurallyUnsafeResourceNames(objects, decodedStreams, placement);
        for (int i = 0; i < decodedStreams.Count; i++) {
            if (InvokesImageInsideMarkedContent(decodedStreams[i].Content, structurallyUnsafeNames)) {
                throw new NotSupportedException("Editing an image inside tagged, artifact, or optional marked content is not supported because its structural context cannot be preserved safely.");
            }
        }

        EnsureFormInvocationIsIsolated(objects, decodedStreams, placement);
    }

    private static bool InvokesImageInsideMarkedContent(string content, HashSet<string> resourceNames) {
        int markedContentDepth = 0;
        bool found = false;
        PdfContentStreamInterpreter.Interpret(
            content,
            PdfReadLimits.DefaultMaxContentOperations,
            operation => {
                if (string.Equals(operation.Name, "BDC", StringComparison.Ordinal) ||
                    string.Equals(operation.Name, "BMC", StringComparison.Ordinal)) {
                    markedContentDepth++;
                } else if (string.Equals(operation.Name, "EMC", StringComparison.Ordinal)) {
                    if (markedContentDepth > 0) markedContentDepth--;
                } else if (markedContentDepth > 0 &&
                           string.Equals(operation.Name, "Do", StringComparison.Ordinal) &&
                           operation.Operands.Count > 0 &&
                           operation.Operands[operation.Operands.Count - 1] is string name &&
                           resourceNames.Contains(name)) {
                    found = true;
                }
            });
        return found;
    }

    private static HashSet<string> ReadStructurallyUnsafeResourceNames(
        Dictionary<int, PdfIndirectObject> objects,
        List<(int ObjectNumber, PdfStream Stream, string Content)> decodedStreams,
        PdfImagePlacement placement) {
        var containingForms = new HashSet<int>();
        for (int i = 0; i < decodedStreams.Count; i++) {
            (int objectNumber, PdfStream stream, string content) = decodedStreams[i];
            if (string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal) &&
                ResourceMapsToObject(objects, stream.Dictionary, placement.ResourceName, placement.ObjectNumber) &&
                InvokesResource(content, placement.ResourceName)) containingForms.Add(objectNumber);
        }

        bool changed;
        do {
            changed = false;
            for (int i = 0; i < decodedStreams.Count; i++) {
                (int objectNumber, PdfStream stream, string content) = decodedStreams[i];
                if (containingForms.Contains(objectNumber) ||
                    !string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) continue;
                foreach (string name in ReadInvokedResourceNames(content)) {
                    if (!ResourceMapsToAnyObject(objects, stream.Dictionary, name, containingForms)) continue;
                    changed = containingForms.Add(objectNumber) || changed;
                    break;
                }
            }
        } while (changed);

        var names = new HashSet<string>(StringComparer.Ordinal) { placement.ResourceName };
        foreach (PdfIndirectObject indirect in objects.Values) {
            PdfDictionary? dictionary = indirect.Value is PdfDictionary direct
                ? direct
                : indirect.Value is PdfStream stream ? stream.Dictionary : null;
            if (dictionary == null) continue;
            AddResourceNamesMappingToObjects(objects, dictionary, containingForms, names);
        }
        return names;
    }

    private static void EnsureFormInvocationIsIsolated(
        Dictionary<int, PdfIndirectObject> objects,
        List<(int ObjectNumber, PdfStream Stream, string Content)> decodedStreams,
        PdfImagePlacement placement) {
        var containingForms = new HashSet<int>();
        for (int i = 0; i < decodedStreams.Count; i++) {
            (int objectNumber, PdfStream stream, string content) = decodedStreams[i];
            if (!string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal) ||
                !ResourceMapsToObject(objects, stream.Dictionary, placement.ResourceName, placement.ObjectNumber) ||
                !InvokesResource(content, placement.ResourceName)) continue;
            containingForms.Add(objectNumber);
        }
        if (containingForms.Count == 0) return;

        var streamsByForm = new Dictionary<int, HashSet<int>>();
        for (int i = 0; i < decodedStreams.Count; i++) {
            foreach (string name in ReadInvokedResourceNames(decodedStreams[i].Content)) {
                foreach (int formObjectNumber in containingForms) {
                    if (!AnyResourceDictionaryMapsToObject(objects, name, formObjectNumber)) continue;
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
        PdfDictionary? resources = owner.Items.TryGetValue("Resources", out PdfObject? resourcesObject)
            ? PdfObjectLookup.Resolve(objects, resourcesObject) as PdfDictionary
            : null;
        PdfDictionary? xObjects = resources?.Items.TryGetValue("XObject", out PdfObject? xObjectsObject) == true
            ? PdfObjectLookup.Resolve(objects, xObjectsObject) as PdfDictionary
            : null;
        return xObjects?.Items.TryGetValue(resourceName, out PdfObject? value) == true &&
               value is PdfReference reference && reference.ObjectNumber == objectNumber;
    }

    private static bool AnyResourceDictionaryMapsToObject(
        Dictionary<int, PdfIndirectObject> objects,
        string resourceName,
        int objectNumber) {
        foreach (PdfIndirectObject indirect in objects.Values) {
            PdfDictionary? dictionary = indirect.Value is PdfDictionary direct
                ? direct
                : indirect.Value is PdfStream stream ? stream.Dictionary : null;
            if (dictionary is not null && ResourceMapsToObject(objects, dictionary, resourceName, objectNumber)) return true;
            if (dictionary?.Items.TryGetValue("XObject", out PdfObject? xObjectsObject) == true &&
                PdfObjectLookup.Resolve(objects, xObjectsObject) is PdfDictionary xObjects &&
                xObjects.Items.TryGetValue(resourceName, out PdfObject? value) &&
                value is PdfReference reference && reference.ObjectNumber == objectNumber) return true;
        }
        return false;
    }

    private static bool ResourceMapsToAnyObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary owner,
        string resourceName,
        HashSet<int> objectNumbers) {
        PdfDictionary? resources = owner.Items.TryGetValue("Resources", out PdfObject? resourcesObject)
            ? PdfObjectLookup.Resolve(objects, resourcesObject) as PdfDictionary
            : null;
        PdfDictionary? xObjects = resources?.Items.TryGetValue("XObject", out PdfObject? xObjectsObject) == true
            ? PdfObjectLookup.Resolve(objects, xObjectsObject) as PdfDictionary
            : null;
        return xObjects?.Items.TryGetValue(resourceName, out PdfObject? value) == true &&
               value is PdfReference reference && objectNumbers.Contains(reference.ObjectNumber);
    }

    private static void AddResourceNamesMappingToObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary owner,
        HashSet<int> objectNumbers,
        HashSet<string> names) {
        PdfDictionary? resources = owner.Items.TryGetValue("Resources", out PdfObject? resourcesObject)
            ? PdfObjectLookup.Resolve(objects, resourcesObject) as PdfDictionary
            : owner;
        if (resources == null) return;
        PdfDictionary? xObjects = resources.Items.TryGetValue("XObject", out PdfObject? xObjectsObject)
            ? PdfObjectLookup.Resolve(objects, xObjectsObject) as PdfDictionary
            : null;
        if (xObjects == null) return;
        foreach (KeyValuePair<string, PdfObject> item in xObjects.Items) {
            if (item.Value is PdfReference reference && objectNumbers.Contains(reference.ObjectNumber)) names.Add(item.Key);
        }
    }

    private static bool InvokesResource(string content, string resourceName) =>
        ReadInvokedResourceNames(content).Any(name => string.Equals(name, resourceName, StringComparison.Ordinal));

    private static List<string> ReadInvokedResourceNames(string content) {
        var names = new List<string>();
        PdfContentStreamInterpreter.Interpret(
            content,
            PdfReadLimits.DefaultMaxContentOperations,
            operation => {
                if (string.Equals(operation.Name, "Do", StringComparison.Ordinal) &&
                    operation.Operands.Count > 0 &&
                    operation.Operands[operation.Operands.Count - 1] is string name) names.Add(name);
            });
        return names;
    }
}

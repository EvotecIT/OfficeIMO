namespace OfficeIMO.Pdf;

internal static class PdfPageResourceHelper {
    internal static PdfDictionary EnsurePageXObjects(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, string operationName) {
        PdfDictionary resources;
        if (page.Items.TryGetValue("Resources", out var resourcesObject)) {
            resources = ResolveDictionary(objects, resourcesObject) ?? throw new NotSupportedException("Indirect page resources must resolve to a dictionary before " + operationName + ".");
            if (resourcesObject is PdfReference) {
                resources = CloneDictionary(resources);
                page.Items["Resources"] = resources;
            }
        } else if (TryResolveInheritedResources(objects, page, out var inheritedResources)) {
            resources = CloneDictionary(inheritedResources);
            page.Items["Resources"] = resources;
        } else {
            resources = new PdfDictionary();
            page.Items["Resources"] = resources;
        }

        if (resources.Items.TryGetValue("XObject", out var xObjectObject)) {
            if (ResolveObject(objects, xObjectObject) is PdfDictionary existing) {
                if (xObjectObject is PdfReference) {
                    var clonedXObjects = CloneDictionary(existing);
                    resources.Items["XObject"] = clonedXObjects;
                    return clonedXObjects;
                }

                return existing;
            }

            throw new NotSupportedException("Page XObject resources must be a dictionary before " + operationName + ".");
        }

        var xObjects = new PdfDictionary();
        resources.Items["XObject"] = xObjects;
        return xObjects;
    }

    private static bool TryResolveInheritedResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, out PdfDictionary resources) {
        resources = null!;
        PdfDictionary? current = ResolveParentDictionary(objects, page);
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.TryGetValue("Resources", out var resourcesObject) &&
                ResolveDictionary(objects, resourcesObject) is PdfDictionary resolved) {
                resources = resolved;
                return true;
            }

            current = ResolveParentDictionary(objects, current);
        }

        return false;
    }

    private static PdfDictionary? ResolveParentDictionary(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary) {
        return dictionary.Items.TryGetValue("Parent", out var parentObject) &&
            parentObject is PdfReference parentReference &&
            PdfObjectLookup.TryGet(objects, parentReference, out var parentIndirect) &&
            parentIndirect.Value is PdfDictionary parentDictionary
            ? parentDictionary
            : null;
    }

    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        return PdfObjectLookup.Resolve(objects, value);
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        return ResolveObject(objects, value) as PdfDictionary;
    }

    private static PdfDictionary CloneDictionary(PdfDictionary dictionary) {
        var clone = new PdfDictionary();
        foreach (var entry in dictionary.Items) {
            clone.Items[entry.Key] = CloneObject(entry.Value);
        }

        return clone;
    }

    private static PdfObject CloneObject(PdfObject value) {
        switch (value) {
            case PdfDictionary dictionary:
                return CloneDictionary(dictionary);
            case PdfArray array:
                var arrayClone = new PdfArray();
                foreach (var item in array.Items) {
                    arrayClone.Items.Add(CloneObject(item));
                }

                return arrayClone;
            case PdfStream stream:
                return new PdfStream(CloneDictionary(stream.Dictionary), (byte[])stream.Data.Clone(), stream.DecodingFailed, stream.DecodingError);
            default:
                return value;
        }
    }
}

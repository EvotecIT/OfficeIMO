using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    private static Dictionary<string, PdfObject> BuildPageOverrides(
        Dictionary<int, PdfIndirectObject> objects,
        int pageObjectNumber,
        string fontResourceName,
        int stampPseudoObjectNumber,
        bool behindContent) {
        if (!objects.TryGetValue(pageObjectNumber, out var indirect) || indirect.Value is not PdfDictionary pageDictionary) {
            throw new InvalidOperationException("PDF page object " + pageObjectNumber.ToString(CultureInfo.InvariantCulture) + " was not found.");
        }

        var contents = BuildContentsArray(objects, pageDictionary.Items.TryGetValue("Contents", out var contentsObj) ? contentsObj : null, stampPseudoObjectNumber, behindContent);
        var resources = BuildResourcesDictionary(objects, GetInheritedPageValue(objects, pageDictionary, "Resources"), fontResourceName);

        return new Dictionary<string, PdfObject>(StringComparer.Ordinal) {
            ["Contents"] = contents,
            ["Resources"] = resources
        };
    }

    private static Dictionary<string, PdfObject> BuildImagePageOverrides(
        Dictionary<int, PdfIndirectObject> objects,
        int pageObjectNumber,
        string imageResourceName,
        int stampPseudoObjectNumber,
        bool behindContent) {
        if (!objects.TryGetValue(pageObjectNumber, out var indirect) || indirect.Value is not PdfDictionary pageDictionary) {
            throw new InvalidOperationException("PDF page object " + pageObjectNumber.ToString(CultureInfo.InvariantCulture) + " was not found.");
        }

        var contents = BuildContentsArray(objects, pageDictionary.Items.TryGetValue("Contents", out var contentsObj) ? contentsObj : null, stampPseudoObjectNumber, behindContent);
        var resources = BuildImageResourcesDictionary(objects, GetInheritedPageValue(objects, pageDictionary, "Resources"), imageResourceName);

        return new Dictionary<string, PdfObject>(StringComparer.Ordinal) {
            ["Contents"] = contents,
            ["Resources"] = resources
        };
    }

    private static PdfArray BuildContentsArray(Dictionary<int, PdfIndirectObject> objects, PdfObject? existingContents, int stampPseudoObjectNumber, bool behindContent) {
        var result = new PdfArray();
        var stampReference = new PdfReference(stampPseudoObjectNumber, 0);

        if (behindContent) {
            result.Items.Add(stampReference);
        }

        AppendContentEntries(objects, result, existingContents);

        if (!behindContent) {
            result.Items.Add(stampReference);
        }

        return result;
    }

    private static void AppendContentEntries(Dictionary<int, PdfIndirectObject> objects, PdfArray target, PdfObject? contents) {
        if (contents is null) {
            return;
        }

        if (contents is PdfArray directArray) {
            foreach (var item in directArray.Items) {
                target.Items.Add(item);
            }

            return;
        }

        if (contents is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            foreach (var item in referencedArray.Items) {
                target.Items.Add(item);
            }

            return;
        }

        target.Items.Add(contents);
    }

    private static PdfDictionary BuildResourcesDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? existingResources,
        string fontResourceName) {
        var resources = CloneDictionary(ResolveDictionary(objects, existingResources));
        var fonts = CloneDictionary(ResolveDictionary(objects, resources.Items.TryGetValue("Font", out var fontObj) ? fontObj : null));
        fonts.Items[fontResourceName] = new PdfReference(FontPseudoObjectNumber, 0);
        resources.Items["Font"] = fonts;
        return resources;
    }

    private static PdfDictionary BuildImageResourcesDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? existingResources,
        string imageResourceName) {
        var resources = CloneDictionary(ResolveDictionary(objects, existingResources));
        var xObjects = CloneDictionary(ResolveDictionary(objects, resources.Items.TryGetValue("XObject", out var xObjectObj) ? xObjectObj : null));
        xObjects.Items[imageResourceName] = new PdfReference(ImagePseudoObjectNumber, 0);
        resources.Items["XObject"] = xObjects;
        return resources;
    }

    private static PdfDictionary CloneDictionary(PdfDictionary? source) {
        var clone = new PdfDictionary();
        if (source is null) {
            return clone;
        }

        foreach (var entry in source.Items) {
            clone.Items[entry.Key] = entry.Value;
        }

        return clone;
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? obj) {
        if (obj is PdfDictionary dictionary) {
            return dictionary;
        }

        if (obj is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
            indirect.Value is PdfDictionary referencedDictionary) {
            return referencedDictionary;
        }

        return null;
    }

    private static string GetAvailableFontResourceName(Dictionary<int, PdfIndirectObject> objects, int[] pageObjectNumbers) {
        var usedNames = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < pageObjectNumbers.Length; i++) {
            if (!objects.TryGetValue(pageObjectNumbers[i], out var indirect) || indirect.Value is not PdfDictionary pageDictionary) {
                continue;
            }

            var resources = ResolveDictionary(objects, GetInheritedPageValue(objects, pageDictionary, "Resources"));
            var fonts = ResolveDictionary(objects, resources?.Items.TryGetValue("Font", out var fontObj) == true ? fontObj : null);
            if (fonts is null) {
                continue;
            }

            foreach (string name in fonts.Items.Keys) {
                usedNames.Add(name);
            }
        }

        const string baseName = "OIMOStampF";
        for (int i = 1; i < 1000; i++) {
            string candidate = baseName + i.ToString(CultureInfo.InvariantCulture);
            if (!usedNames.Contains(candidate)) {
                return candidate;
            }
        }

        throw new InvalidOperationException("Unable to find an available PDF font resource name for the stamp.");
    }

    private static string GetAvailableXObjectResourceName(
        Dictionary<int, PdfIndirectObject> objects,
        int[] pageObjectNumbers,
        HashSet<string>? additionallyUsed = null) {
        var usedNames = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < pageObjectNumbers.Length; i++) {
            if (!objects.TryGetValue(pageObjectNumbers[i], out var indirect) || indirect.Value is not PdfDictionary pageDictionary) {
                continue;
            }

            var resources = ResolveDictionary(objects, GetInheritedPageValue(objects, pageDictionary, "Resources"));
            var xObjects = ResolveDictionary(objects, resources?.Items.TryGetValue("XObject", out var xObjectObj) == true ? xObjectObj : null);
            if (xObjects is null) {
                continue;
            }

            foreach (string name in xObjects.Items.Keys) {
                usedNames.Add(name);
            }
        }

        const string baseName = "OIMOStampIm";
        for (int i = 1; i < 1000; i++) {
            string candidate = baseName + i.ToString(CultureInfo.InvariantCulture);
            if (!usedNames.Contains(candidate) && (additionallyUsed is null || !additionallyUsed.Contains(candidate))) {
                return candidate;
            }
        }

        throw new InvalidOperationException("Unable to find an available PDF image resource name for the stamp.");
    }

    private static PdfObject? GetInheritedPageValue(Dictionary<int, PdfIndirectObject> objects, PdfDictionary pageDictionary, string key) {
        PdfDictionary? current = pageDictionary;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.TryGetValue(key, out var value)) {
                return value;
            }

            if (!current.Items.TryGetValue("Parent", out var parentObj) ||
                parentObj is not PdfReference parentReference ||
                !PdfObjectLookup.TryGet(objects, parentReference, out var parentIndirect) ||
                parentIndirect.Value is not PdfDictionary parentDictionary) {
                return null;
            }

            current = parentDictionary;
        }

        return null;
    }

    private static PdfDictionary BuildFontObject(PdfStandardFont font) {
        return PdfStandardFontDictionaryBuilder.BuildStandardType1FontDictionary(font);
    }

}

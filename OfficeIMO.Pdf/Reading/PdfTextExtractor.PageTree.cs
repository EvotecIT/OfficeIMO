using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

public static partial class PdfTextExtractor {
    private static List<PdfDictionary> CollectPages(Dictionary<int, PdfIndirectObject> objects, string? trailerRaw) {
        var result = new List<PdfDictionary>();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, trailerRaw);
        if (catalog is not null &&
            ResolveDict(catalog.Items.TryGetValue("Pages", out var pagesObj) ? pagesObj : null, objects) is PdfDictionary pagesRoot) {
            TraversePagesNode(pagesRoot, objects, result, new HashSet<int>());
        }
    
        if (result.Count > 0) {
            return result;
        }
    
        foreach (var kv in objects.OrderBy(k => k.Key)) {
            if (kv.Value.Value is PdfDictionary dict && IsLeafPage(dict, objects)) {
                result.Add(dict);
            }
        }
    
        return result;
    }
    
    private static void TraversePagesNode(
        PdfDictionary node,
        Dictionary<int, PdfIndirectObject> objects,
        List<PdfDictionary> result,
        HashSet<int> visited) {
        int objectNumber = FindObjectNumberFor(node, objects);
        if (objectNumber > 0 && !visited.Add(objectNumber)) {
            return;
        }
    
        string? type = node.Get<PdfName>("Type")?.Name;
        if (type == "Page" || (type is null && IsLeafPage(node, objects))) {
            result.Add(node);
            return;
        }
    
        var kids = ResolveArray(node.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null, objects);
        if (kids is null) {
            return;
        }
    
        foreach (var kid in kids.Items) {
            var child = ResolveDict(kid, objects);
            if (child is not null) {
                TraversePagesNode(child, objects, result, visited);
            }
        }
    }
    
    private static bool IsLeafPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        if (ResolveArray(page.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null, objects) is not null) {
            return false;
        }
    
        if (!page.Items.ContainsKey("Contents")) {
            return false;
        }
    
        string? type = page.Get<PdfName>("Type")?.Name;
        if (type == "Page") {
            return true;
        }
    
        return type is null &&
               (page.Items.ContainsKey("Resources") || GetInheritedValue(page, "Resources", objects) is not null) &&
               (page.Items.ContainsKey("MediaBox") ||
                page.Items.ContainsKey("CropBox") ||
                GetInheritedValue(page, "MediaBox", objects) is not null ||
                GetInheritedValue(page, "CropBox", objects) is not null);
    }
    
    private static List<int> GetContentIds(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var ids = new List<int>();
        if (!page.Items.TryGetValue("Contents", out var contents)) {
            return ids;
        }
    
        if (contents is PdfReference reference) {
            if (PdfObjectLookup.TryGet(objects, reference, out var indirect)) {
                if (indirect.Value is PdfArray referencedArray) {
                    AppendContentIds(referencedArray, ids, objects);
                } else if (indirect.Value is PdfStream) {
                    ids.Add(reference.ObjectNumber);
                }
            }
            return ids;
        }
    
        if (contents is PdfArray arr) {
            AppendContentIds(arr, ids, objects);
        }
    
        return ids;
    }
    
    private static PdfObject? GetInheritedValue(PdfDictionary start, string key, Dictionary<int, PdfIndirectObject> objects) {
        PdfDictionary? current = start;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.TryGetValue(key, out var value)) {
                return value;
            }
    
            if (!current.Items.TryGetValue("Parent", out var parentObj)) {
                break;
            }
    
            current = ResolveDict(parentObj, objects);
        }
    
        return null;
    }
    
    private static PdfDictionary? ResolveDict(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is PdfDictionary dict) {
            return dict;
        }
    
        if (obj is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
            indirect.Value is PdfDictionary referencedDict) {
            return referencedDict;
        }
    
        return null;
    }
    
    private static PdfArray? ResolveArray(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is PdfArray arr) {
            return arr;
        }
    
        if (obj is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            return referencedArray;
        }
    
        return null;
    }
    
    private static void AppendContentIds(PdfArray contentArray, List<int> ids, Dictionary<int, PdfIndirectObject> objects) {
        foreach (var item in contentArray.Items) {
            if (item is PdfReference itemReference &&
                PdfObjectLookup.TryGet(objects, itemReference, out var indirect) &&
                indirect.Value is PdfStream) {
                ids.Add(itemReference.ObjectNumber);
            }
        }
    }
    
    private static int FindObjectNumberFor(PdfDictionary dict, Dictionary<int, PdfIndirectObject> objects) {
        foreach (var kv in objects) {
            if (ReferenceEquals(kv.Value.Value, dict)) {
                return kv.Key;
            }
        }

        return 0;
    }
}

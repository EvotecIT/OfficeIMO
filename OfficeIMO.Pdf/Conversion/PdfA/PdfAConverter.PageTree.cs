using System.Globalization;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    private static readonly string[] InheritedPageKeys = { "Resources", "MediaBox", "CropBox", "Rotate" };

    /// <summary>
    /// Validators reject page trees with loops or wrong counts that readers tolerate. When the tree does not list exactly the
    /// pages that were read, in order, it is rebuilt as one flat list of those pages, with inherited attributes pushed down.
    /// </summary>
    private void RepairPageTree(PdfDictionary catalog) {
        if (catalog.Items.TryGetValue("Pages", out PdfObject? rootValue) is false || rootValue is not PdfReference rootReference ||
            ResolveDictionary(rootValue) is not PdfDictionary root) return;
        IReadOnlyList<PdfReadPage> pages = _document.Pages;
        if (pages.Count == 0) {
            AddViolation(PdfAConversionRefusalReason.SourceUnreadable, "page tree", "No page could be read from the page tree.", root);
            return;
        }
        var leaves = new List<int>();
        bool consistent = CollectLeaves(root, new HashSet<PdfObject>(ReferenceComparer.Instance), leaves, 0, out int declared) &&
            declared == pages.Count && leaves.Count == pages.Count;
        var readPages = new HashSet<int>(pages.Select(static page => page.ObjectNumber));
        foreach (int leaf in leaves) {
            if (!readPages.Contains(leaf)) {
                AddViolation(PdfAConversionRefusalReason.SourceUnreadable, "page tree",
                    "The page tree lists a page that could not be read; rebuilding the tree would drop it.", leaf);
                return;
            }
        }
        for (int index = 0; consistent && index < pages.Count; index++) consistent = leaves[index] == pages[index].ObjectNumber;
        if (consistent) return;
        if (!Change(PdfAConversionChangeKind.StructureUpdated, "page tree",
                "Rebuilt the page tree, which had a loop or inconsistent counts, as one list of the same " +
                pages.Count.ToString(CultureInfo.InvariantCulture) + " page(s) in the same order.", root)) return;

        var kids = new PdfArray();
        foreach (PdfReadPage page in pages) {
            PdfDictionary dictionary = page.PageDictionary;
            foreach (string key in InheritedPageKeys) {
                if (dictionary.Items.ContainsKey(key)) continue;
                PdfObject? inherited = FindInherited(dictionary, key);
                if (inherited is not null) dictionary.Items[key] = inherited;
            }
            dictionary.Items["Parent"] = rootReference;
            int generation = _objects.TryGetValue(page.ObjectNumber, out PdfIndirectObject? indirect) ? indirect.Generation : 0;
            kids.Items.Add(new PdfReference(page.ObjectNumber, generation));
        }
        root.Items["Kids"] = kids;
        root.Items["Count"] = new PdfNumber(pages.Count);
        root.Items.Remove("Parent");
    }

    /// <summary>Collects every distinct leaf, continuing past loops; returns false when the tree is malformed.</summary>
    private bool CollectLeaves(PdfDictionary node, HashSet<PdfObject> visited, List<int> leaves, int depth, out int declared) {
        declared = Resolve(node.Items.TryGetValue("Count", out PdfObject? count) ? count : null) is PdfNumber number ? (int)number.Value : -1;
        if (depth > 256 || !visited.Add(node)) return false;
        if (Resolve(node.Items.TryGetValue("Kids", out PdfObject? kids) ? kids : null) is not PdfArray array) return false;
        bool wellFormed = true;
        foreach (PdfObject kid in array.Items) {
            if (kid is not PdfReference reference || ResolveDictionary(kid) is not PdfDictionary child) {
                wellFormed = false;
                continue;
            }
            if (GetName(child, "Type") == "Pages" || child.Items.ContainsKey("Kids")) {
                wellFormed &= CollectLeaves(child, visited, leaves, depth + 1, out _);
            } else if (visited.Add(child)) {
                leaves.Add(reference.ObjectNumber);
            } else {
                wellFormed = false;
            }
        }
        return wellFormed;
    }

    private PdfObject? FindInherited(PdfDictionary page, string key) {
        var visited = new HashSet<PdfObject>(ReferenceComparer.Instance);
        PdfDictionary? current = ResolveDictionary(page.Items.TryGetValue("Parent", out PdfObject? parent) ? parent : null);
        while (current is not null && visited.Add(current)) {
            if (current.Items.TryGetValue(key, out PdfObject? value)) return value;
            current = ResolveDictionary(current.Items.TryGetValue("Parent", out PdfObject? next) ? next : null);
        }
        return null;
    }
}

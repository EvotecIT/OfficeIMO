namespace OfficeIMO.Pdf;

/// <summary>Builds the deterministic object ownership and ordering plan required by a linearized PDF.</summary>
internal static class PdfLinearizationPlanner {
    private static readonly string[] InheritedPageKeys = {
        "Resources", "MediaBox", "CropBox", "Rotate", "BleedBox", "TrimBox", "ArtBox"
    };

    internal static PdfLinearizationPlan Create(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber) {
        Guard.NotNull(objects, nameof(objects));
        if (!objects.TryGetValue(catalogObjectNumber, out PdfIndirectObject? catalogObject) ||
            catalogObject.Value is not PdfDictionary catalog) {
            throw new ArgumentException("PDF catalog object is unavailable for linearization.", nameof(catalogObjectNumber));
        }

        var pageObjectNumbers = new List<int>();
        var pageTreeObjectNumbers = new HashSet<int>();
        if (!catalog.Items.TryGetValue("Pages", out PdfObject? pagesRoot)) {
            throw new InvalidOperationException("PDF catalog does not contain a page tree.");
        }

        CollectPages(pagesRoot, objects, pageObjectNumbers, pageTreeObjectNumbers, new HashSet<int>());
        if (pageObjectNumbers.Count == 0) throw new InvalidOperationException("PDF does not contain any indirect page objects to linearize.");

        var pageObjectSet = new HashSet<int>(pageObjectNumbers);
        for (int index = 0; index < pageObjectNumbers.Count; index++) {
            MaterializeInheritedPageValues(pageObjectNumbers[index], objects);
        }

        var reachableByPage = new List<List<int>>(pageObjectNumbers.Count);
        var reachableSets = new List<HashSet<int>>(pageObjectNumbers.Count);
        for (int index = 0; index < pageObjectNumbers.Count; index++) {
            int pageObjectNumber = pageObjectNumbers[index];
            var ordered = new List<int> { pageObjectNumber };
            var seen = new HashSet<int> { pageObjectNumber };
            CollectPageReachableObjects(
                objects[pageObjectNumber].Value,
                objects,
                pageObjectSet,
                pageTreeObjectNumbers,
                catalogObjectNumber,
                seen,
                ordered,
                skipParentEntry: true);
            reachableByPage.Add(ordered);
            reachableSets.Add(seen);
        }

        var firstPageSet = new HashSet<int>(reachableSets[0]);
        var nonFirstReferenceCounts = new Dictionary<int, int>();
        for (int pageIndex = 1; pageIndex < reachableSets.Count; pageIndex++) {
            foreach (int objectNumber in reachableSets[pageIndex]) {
                if (objectNumber == pageObjectNumbers[pageIndex] || firstPageSet.Contains(objectNumber) ||
                    pageObjectSet.Contains(objectNumber) || pageTreeObjectNumbers.Contains(objectNumber) ||
                    objectNumber == catalogObjectNumber) continue;
                nonFirstReferenceCounts.TryGetValue(objectNumber, out int count);
                nonFirstReferenceCounts[objectNumber] = count + 1;
            }
        }

        var sharedObjectSet = new HashSet<int>(nonFirstReferenceCounts.Where(static entry => entry.Value > 1).Select(static entry => entry.Key));
        var assigned = new HashSet<int>(firstPageSet) { catalogObjectNumber };
        var pageGroups = new List<IReadOnlyList<int>>(pageObjectNumbers.Count) { reachableByPage[0].AsReadOnly() };
        for (int pageIndex = 1; pageIndex < reachableByPage.Count; pageIndex++) {
            int pageObjectNumber = pageObjectNumbers[pageIndex];
            var group = new List<int> { pageObjectNumber };
            assigned.Add(pageObjectNumber);
            foreach (int objectNumber in reachableByPage[pageIndex]) {
                if (objectNumber == pageObjectNumber || firstPageSet.Contains(objectNumber) || sharedObjectSet.Contains(objectNumber) ||
                    pageObjectSet.Contains(objectNumber) || pageTreeObjectNumbers.Contains(objectNumber) || objectNumber == catalogObjectNumber ||
                    assigned.Contains(objectNumber)) continue;
                group.Add(objectNumber);
                assigned.Add(objectNumber);
            }

            pageGroups.Add(group.AsReadOnly());
        }

        var sharedObjects = new List<int>();
        for (int pageIndex = 1; pageIndex < reachableByPage.Count; pageIndex++) {
            foreach (int objectNumber in reachableByPage[pageIndex]) {
                if (!sharedObjectSet.Contains(objectNumber) || assigned.Contains(objectNumber)) continue;
                sharedObjects.Add(objectNumber);
                assigned.Add(objectNumber);
            }
        }

        var remainingObjects = objects.Keys
            .Where(objectNumber => !assigned.Contains(objectNumber) && objectNumber != catalogObjectNumber)
            .OrderBy(static objectNumber => objectNumber)
            .ToList();

        return new PdfLinearizationPlan(
            catalogObjectNumber,
            pageObjectNumbers.AsReadOnly(),
            pageGroups.AsReadOnly(),
            reachableSets.AsReadOnly(),
            sharedObjects.AsReadOnly(),
            remainingObjects.AsReadOnly());
    }

    private static void CollectPages(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects,
        List<int> pageObjectNumbers,
        HashSet<int> pageTreeObjectNumbers,
        HashSet<int> visited) {
        if (value is not PdfReference reference || !objects.TryGetValue(reference.ObjectNumber, out PdfIndirectObject? indirect) || !visited.Add(reference.ObjectNumber)) return;
        if (indirect.Value is not PdfDictionary dictionary) return;

        string? type = dictionary.Get<PdfName>("Type")?.Name;
        if (string.Equals(type, "Page", StringComparison.Ordinal)) {
            pageObjectNumbers.Add(reference.ObjectNumber);
            return;
        }

        if (!string.Equals(type, "Pages", StringComparison.Ordinal) && !dictionary.Items.ContainsKey("Kids")) return;
        pageTreeObjectNumbers.Add(reference.ObjectNumber);
        if (!dictionary.Items.TryGetValue("Kids", out PdfObject? kidsObject) || PdfObjectLookup.Resolve(objects, kidsObject) is not PdfArray kids) return;
        foreach (PdfObject kid in kids.Items) CollectPages(kid, objects, pageObjectNumbers, pageTreeObjectNumbers, visited);
    }

    private static void MaterializeInheritedPageValues(int pageObjectNumber, Dictionary<int, PdfIndirectObject> objects) {
        if (!objects.TryGetValue(pageObjectNumber, out PdfIndirectObject? pageObject) || pageObject.Value is not PdfDictionary page) return;
        PdfDictionary? ancestor = ResolveParent(page, objects);
        var visitedParents = new HashSet<PdfDictionary>();
        while (ancestor != null && visitedParents.Add(ancestor)) {
            for (int index = 0; index < InheritedPageKeys.Length; index++) {
                string key = InheritedPageKeys[index];
                if (!page.Items.ContainsKey(key) && ancestor.Items.TryGetValue(key, out PdfObject? value)) page.Items[key] = value;
            }

            ancestor = ResolveParent(ancestor, objects);
        }
    }

    private static PdfDictionary? ResolveParent(PdfDictionary dictionary, Dictionary<int, PdfIndirectObject> objects) {
        if (!dictionary.Items.TryGetValue("Parent", out PdfObject? parent)) return null;
        return PdfObjectLookup.Resolve(objects, parent) as PdfDictionary;
    }

    private static void CollectPageReachableObjects(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> pageObjectNumbers,
        HashSet<int> pageTreeObjectNumbers,
        int catalogObjectNumber,
        HashSet<int> seen,
        List<int> ordered,
        bool skipParentEntry = false) {
        if (value is PdfReference reference) {
            int objectNumber = reference.ObjectNumber;
            if (objectNumber == catalogObjectNumber || pageTreeObjectNumbers.Contains(objectNumber) || pageObjectNumbers.Contains(objectNumber) ||
                !objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) || !seen.Add(objectNumber)) return;
            ordered.Add(objectNumber);
            CollectPageReachableObjects(indirect.Value, objects, pageObjectNumbers, pageTreeObjectNumbers, catalogObjectNumber, seen, ordered);
            return;
        }

        if (value is PdfArray array) {
            foreach (PdfObject item in array.Items) CollectPageReachableObjects(item, objects, pageObjectNumbers, pageTreeObjectNumbers, catalogObjectNumber, seen, ordered);
            return;
        }

        PdfDictionary? dictionary = value switch {
            PdfDictionary directDictionary => directDictionary,
            PdfStream stream => stream.Dictionary,
            _ => null
        };
        if (dictionary == null) return;
        foreach (KeyValuePair<string, PdfObject> entry in dictionary.Items) {
            if (skipParentEntry && string.Equals(entry.Key, "Parent", StringComparison.Ordinal)) continue;
            CollectPageReachableObjects(entry.Value, objects, pageObjectNumbers, pageTreeObjectNumbers, catalogObjectNumber, seen, ordered);
        }
    }
}

internal sealed class PdfLinearizationPlan {
    internal PdfLinearizationPlan(
        int catalogObjectNumber,
        IReadOnlyList<int> pageObjectNumbers,
        IReadOnlyList<IReadOnlyList<int>> pageGroups,
        IReadOnlyList<HashSet<int>> reachableByPage,
        IReadOnlyList<int> sharedObjects,
        IReadOnlyList<int> remainingObjects) {
        CatalogObjectNumber = catalogObjectNumber;
        PageObjectNumbers = pageObjectNumbers;
        PageGroups = pageGroups;
        ReachableByPage = reachableByPage;
        SharedObjects = sharedObjects;
        RemainingObjects = remainingObjects;
    }

    internal int CatalogObjectNumber { get; }
    internal IReadOnlyList<int> PageObjectNumbers { get; }
    internal IReadOnlyList<IReadOnlyList<int>> PageGroups { get; }
    internal IReadOnlyList<HashSet<int>> ReachableByPage { get; }
    internal IReadOnlyList<int> SharedObjects { get; }
    internal IReadOnlyList<int> RemainingObjects { get; }
}

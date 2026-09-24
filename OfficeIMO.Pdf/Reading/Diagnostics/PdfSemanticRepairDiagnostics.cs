namespace OfficeIMO.Pdf;

internal static class PdfSemanticRepairDiagnostics {
    internal static IReadOnlyList<PdfRepairDiagnostic> AnalyzeAndRepair(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary? catalog,
        IReadOnlyList<PdfReadPage> pages,
        PdfLoadOptions options,
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var diagnostics = new List<PdfRepairDiagnostic>();
        if (catalog is null) return diagnostics;
        PdfDictionary? pagesRoot = ResolveDictionary(objects, catalog.Items.TryGetValue("Pages", out PdfObject? pagesObject) ? pagesObject : null);
        if (pagesRoot is null) {
            AddDefect(options, diagnostics, "BrokenPageTreeRoot", "The catalog /Pages reference does not resolve to a page-tree dictionary; lenient loading used bounded page-object discovery.", null, recovered: pages.Count > 0);
        } else {
            RepairPageTree(objects, pagesRoot, pages, options, diagnostics, cancellationToken);
        }

        ValidateNameTrees(objects, catalog, pages, options, diagnostics, cancellationToken);
        DiagnoseOrphanedSemanticObjects(objects, catalog, diagnostics, cancellationToken);
        return diagnostics.AsReadOnly();
    }

    private static void RepairPageTree(Dictionary<int, PdfIndirectObject> objects, PdfDictionary root, IReadOnlyList<PdfReadPage> pages, PdfLoadOptions options, List<PdfRepairDiagnostic> diagnostics, System.Threading.CancellationToken cancellationToken) {
        int declared = (int)(root.Get<PdfNumber>("Count")?.Value ?? -1);
        if (declared != pages.Count) {
            AddDefect(options, diagnostics, "IncorrectPageTreeCount", "The page-tree root declares /Count " + declared + " but traversal found " + pages.Count + " page(s); lenient loading rebuilt /Count from reachable leaves.", FindObjectNumber(objects, root, cancellationToken), recovered: true);
            root.Items["Count"] = new PdfNumber(pages.Count);
        }

        if (ResolveArray(objects, root.Items.TryGetValue("Kids", out PdfObject? kidsObject) ? kidsObject : null) is not PdfArray kids) {
            AddDefect(options, diagnostics, "MissingPageTreeKids", "The page-tree root has no readable /Kids array; lenient loading retained pages recovered through parent-chain discovery.", FindObjectNumber(objects, root, cancellationToken), recovered: pages.Count > 0);
            return;
        }

        var visited = new HashSet<PdfDictionary>();
        RepairPageNode(objects, root, kids, options, diagnostics, visited, 1, cancellationToken);
    }

    private static void RepairPageNode(Dictionary<int, PdfIndirectObject> objects, PdfDictionary parent, PdfArray kids, PdfLoadOptions options, List<PdfRepairDiagnostic> diagnostics, HashSet<PdfDictionary> visited, int depth, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (depth > options.Limits.MaxPageTreeDepth || !visited.Add(parent)) return;
        for (int index = kids.Items.Count - 1; index >= 0; index--) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject kidObject = kids.Items[index];
            PdfDictionary? kid = ResolveDictionary(objects, kidObject);
            if (kid is null) {
                int? objectNumber = kidObject is PdfReference reference ? reference.ObjectNumber : null;
                AddDefect(options, diagnostics, "InvalidPageTreeKid", "A page-tree /Kids entry does not resolve to a dictionary; lenient loading removed the unusable entry.", objectNumber, recovered: true);
                kids.Items.RemoveAt(index);
                continue;
            }

            if (!kid.Items.TryGetValue("Parent", out PdfObject? currentParent) || !ResolvesTo(objects, currentParent, parent)) {
                int parentObjectNumber = FindObjectNumber(objects, parent, cancellationToken);
                if (parentObjectNumber > 0) {
                    kid.Items["Parent"] = new PdfReference(parentObjectNumber, objects[parentObjectNumber].Generation);
                    AddDefect(options, diagnostics, "BrokenPageParent", "A page-tree child had a missing or incorrect /Parent reference; lenient loading restored the containing page-tree node.", FindObjectNumber(objects, kid, cancellationToken), recovered: true);
                }
            }

            if (ResolveArray(objects, kid.Items.TryGetValue("Kids", out PdfObject? nestedKidsObject) ? nestedKidsObject : null) is PdfArray nestedKids) RepairPageNode(objects, kid, nestedKids, options, diagnostics, visited, depth + 1, cancellationToken);
        }
    }

    private static void ValidateNameTrees(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog, IReadOnlyList<PdfReadPage> pages, PdfLoadOptions options, List<PdfRepairDiagnostic> diagnostics, System.Threading.CancellationToken cancellationToken) {
        PdfDictionary? names = ResolveDictionary(objects, catalog.Items.TryGetValue("Names", out PdfObject? namesObject) ? namesObject : null);
        if (names is null) return;
        foreach (KeyValuePair<string, PdfObject> entry in names.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!IsStandardCatalogNameTree(entry.Key)) continue;
            int traversedNameTreeNodes = 0;
            ValidateNameTreeNode(objects, entry.Value, entry.Key, new HashSet<(int ObjectNumber, int Generation)>(), options, diagnostics, 0, ref traversedNameTreeNodes, cancellationToken);
        }

        if (names.Items.TryGetValue("Dests", out PdfObject? destinations)) {
            int traversedDestinationTreeNodes = 0;
            ValidateDestinationTree(
                objects,
                destinations,
                new HashSet<(int ObjectNumber, int Generation)>(),
                new HashSet<int>(pages.Select(static page => page.ObjectNumber)),
                options,
                diagnostics,
                0,
                ref traversedDestinationTreeNodes,
                cancellationToken);
        }
    }

    private static bool IsStandardCatalogNameTree(string name) {
        return name is "Dests" or "AP" or "JavaScript" or "Pages" or "Templates" or
            "IDS" or "URLS" or "EmbeddedFiles" or "AlternatePresentations" or "Renditions";
    }

    private static void ValidateNameTreeNode(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject nodeObject,
        string treeName,
        HashSet<(int ObjectNumber, int Generation)> visited,
        PdfLoadOptions options,
        List<PdfRepairDiagnostic> diagnostics,
        int depth,
        ref int traversedNodes,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureNameTreeBudget(options.Limits, depth, traversedNodes);
        if (nodeObject is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation))) return;
            EnsureNameTreeBudget(options.Limits, depth, ++traversedNodes);
        }

        PdfDictionary? node = ResolveDictionary(objects, nodeObject);
        if (node is null) { AddDefect(options, diagnostics, "InvalidNameTreeNode", "The " + treeName + " name tree contains a node that does not resolve to a dictionary.", nodeObject is PdfReference missing ? missing.ObjectNumber : null, recovered: false); return; }
        PdfArray? pairs = ResolveArray(objects, node.Items.TryGetValue("Names", out PdfObject? namesObject) ? namesObject : null);
        PdfArray? kids = ResolveArray(objects, node.Items.TryGetValue("Kids", out PdfObject? kidsObject) ? kidsObject : null);
        bool hasNames = pairs is not null;
        bool hasKids = kids is not null;
        if (hasNames && hasKids) AddDefect(options, diagnostics, "MixedNameTreeNode", "The " + treeName + " name-tree node contains both /Names and /Kids and was left unchanged to avoid changing lookup semantics.", FindObjectNumber(objects, node, cancellationToken), recovered: false);
        if (hasNames && pairs!.Items.Count % 2 != 0) AddDefect(options, diagnostics, "OddNameTreePairs", "The " + treeName + " name-tree /Names array has an unmatched key or value and was left unchanged.", FindObjectNumber(objects, node, cancellationToken), recovered: false);
        if (hasKids) foreach (PdfObject kid in kids!.Items) ValidateNameTreeNode(objects, kid, treeName, visited, options, diagnostics, depth + 1, ref traversedNodes, cancellationToken);
    }

    private static void ValidateDestinationTree(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject nodeObject,
        HashSet<(int ObjectNumber, int Generation)> visited,
        HashSet<int> pageObjectNumbers,
        PdfLoadOptions options,
        List<PdfRepairDiagnostic> diagnostics,
        int depth,
        ref int traversedNodes,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureNameTreeBudget(options.Limits, depth, traversedNodes);
        if (nodeObject is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation))) return;
            EnsureNameTreeBudget(options.Limits, depth, ++traversedNodes);
        }

        PdfDictionary? node = ResolveDictionary(objects, nodeObject); if (node is null) return;
        if (ResolveArray(objects, node.Items.TryGetValue("Names", out PdfObject? namesObject) ? namesObject : null) is PdfArray pairs) {
            for (int i = 1; i < pairs.Items.Count; i += 2) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!IsValidDestination(objects, pairs.Items[i], pageObjectNumbers)) AddDefect(options, diagnostics, "BrokenNamedDestination", "A named destination does not resolve to a reachable page and was left unchanged for caller review.", FindObjectNumber(objects, node, cancellationToken), recovered: false);
            }
        }
        if (ResolveArray(objects, node.Items.TryGetValue("Kids", out PdfObject? kidsObject) ? kidsObject : null) is PdfArray kids) foreach (PdfObject kid in kids.Items) ValidateDestinationTree(objects, kid, visited, pageObjectNumbers, options, diagnostics, depth + 1, ref traversedNodes, cancellationToken);
    }

    private static void EnsureNameTreeBudget(PdfReadLimits limits, int depth, int traversedNodes) {
        if (depth > limits.MaxNameTreeDepth) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.NameTreeDepth, limits.MaxNameTreeDepth, depth);
        }

        if (traversedNodes > limits.MaxNameTreeNodes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.NameTreeNodes, limits.MaxNameTreeNodes, traversedNodes);
        }
    }

    private static bool IsValidDestination(Dictionary<int, PdfIndirectObject> objects, PdfObject destinationObject, HashSet<int> pageObjectNumbers) {
        PdfObject? destination = PdfObjectLookup.Resolve(objects, destinationObject);
        if (destination is PdfDictionary dictionary && dictionary.Items.TryGetValue("D", out PdfObject? explicitDestination)) destination = PdfObjectLookup.Resolve(objects, explicitDestination);
        return destination is PdfArray array && array.Items.Count > 0 && array.Items[0] is PdfReference pageReference && pageObjectNumbers.Contains(pageReference.ObjectNumber);
    }

    private static void DiagnoseOrphanedSemanticObjects(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog, List<PdfRepairDiagnostic> diagnostics, System.Threading.CancellationToken cancellationToken) {
        int catalogNumber = FindObjectNumber(objects, catalog, cancellationToken); if (catalogNumber <= 0) return;
        HashSet<int> reachable = PdfCollectionSizing.CreateHashSet<int>(objects.Count, objects.Count);
        TraverseReferences(objects, new PdfReference(catalogNumber, objects[catalogNumber].Generation), reachable, cancellationToken);
        int orphanCount = 0;
        int firstOrphan = int.MaxValue;
        foreach (PdfIndirectObject indirect in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (reachable.Contains(indirect.ObjectNumber) || !IsSemanticObject(indirect.Value)) continue;
            orphanCount++;
            firstOrphan = Math.Min(firstOrphan, indirect.ObjectNumber);
        }

        if (orphanCount == 0) return;
        diagnostics.Add(new PdfRepairDiagnostic("OrphanedSemanticObjects", "Detected " + orphanCount + " unreachable semantic object(s), beginning with object " + firstOrphan + "; they remain available for forensic inspection and are not silently deleted during read.", firstOrphan, PdfRepairDisposition.DetectedOnly));
    }

    private static bool IsSemanticObject(PdfObject value) {
        PdfDictionary? dictionary = value is PdfDictionary direct ? direct : value is PdfStream stream ? stream.Dictionary : null;
        string? type = dictionary?.Get<PdfName>("Type")?.Name;
        string? subtype = dictionary?.Get<PdfName>("Subtype")?.Name;
        return type is "Page" or "Pages" or "Annot" or "Outlines" or "Filespec" || subtype is "Widget" or "Link" or "FileAttachment";
    }

    private static void TraverseReferences(Dictionary<int, PdfIndirectObject> objects, PdfObject root, HashSet<int> reachable, System.Threading.CancellationToken cancellationToken) {
        var pending = new Stack<PdfObject>(); pending.Push(root);
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject value = pending.Pop();
            if (value is PdfReference reference) {
                if (!reachable.Add(reference.ObjectNumber)) continue;
                if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) pending.Push(indirect.Value);
                else reachable.Remove(reference.ObjectNumber); // A bad generation must not hide a later valid reference.
                continue;
            }
            // Indirect references are the only way a parsed PDF object graph can
            // contain cycles. Direct dictionaries and arrays are materialized as
            // bounded trees by the parser, so reference-identity hashing here only
            // repeats work for every container in otherwise valid documents.
            if (value is PdfArray array) {
                for (int i = 0; i < array.Items.Count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    pending.Push(array.Items[i]);
                }
                continue;
            }
            PdfDictionary? dictionary = value is PdfDictionary direct ? direct : value is PdfStream stream ? stream.Dictionary : null;
            if (dictionary != null) {
                foreach (PdfObject item in dictionary.Items.Values) {
                    cancellationToken.ThrowIfCancellationRequested();
                    pending.Push(item);
                }
            }
        }
    }

    private static void AddDefect(PdfLoadOptions options, List<PdfRepairDiagnostic> diagnostics, string code, string message, int? objectNumber, bool recovered) {
        if (options.ParsingMode == PdfParsingMode.Strict) throw new PdfParseException(code, message, objectNumber);
        diagnostics.Add(new PdfRepairDiagnostic(code, message, objectNumber, recovered ? PdfRepairDisposition.Recovered : PdfRepairDisposition.DetectedOnly));
    }

    private static bool ResolvesTo(Dictionary<int, PdfIndirectObject> objects, PdfObject value, PdfDictionary expected) => ReferenceEquals(PdfObjectLookup.Resolve(objects, value), expected);
    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) => PdfObjectLookup.Resolve(objects, value) as PdfDictionary;
    private static PdfArray? ResolveArray(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) => PdfObjectLookup.Resolve(objects, value) as PdfArray;
    private static int FindObjectNumber(Dictionary<int, PdfIndirectObject> objects, PdfObject value, System.Threading.CancellationToken cancellationToken) {
        foreach (PdfIndirectObject indirect in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (ReferenceEquals(indirect.Value, value) || indirect.Value is PdfStream stream && ReferenceEquals(stream.Dictionary, value)) return indirect.ObjectNumber;
        }
        return 0;
    }

}

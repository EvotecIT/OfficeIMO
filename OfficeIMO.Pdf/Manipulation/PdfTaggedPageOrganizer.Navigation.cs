namespace OfficeIMO.Pdf;

internal static partial class PdfTaggedPageOrganizer {
    /// <summary>
    /// Returns the destination names whose every definition goes to a removed page. Names that cannot be resolved,
    /// or that are defined in a name tree this writer cannot traverse, are not included: they are kept as found.
    /// </summary>
    internal static HashSet<string> CollectRemovedDestinationNames(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog,
        HashSet<int> removedPages, CancellationToken token) {
        var survives = new Dictionary<string, bool>(StringComparer.Ordinal);
        if (removedPages.Count == 0) return new HashSet<string>(StringComparer.Ordinal);
        if (catalog.Items.TryGetValue("Dests", out var directValue) && PdfObjectLookup.Resolve(objects, directValue) is PdfDictionary direct)
            foreach (var pair in direct.Items) Record(pair.Key, pair.Value);
        if (TryGetDestinationNameTree(objects, catalog, out _, out var tree)) {
            var pairs = new List<(PdfStringObj Name, PdfObject Value)>();
            if (TryCollectNameTreePairs(objects, tree, pairs, token))
                foreach (var pair in pairs) Record(pair.Name.Value, pair.Value);
        }
        return new HashSet<string>(survives.Where(static pair => !pair.Value).Select(static pair => pair.Key), StringComparer.Ordinal);

        void Record(string name, PdfObject value) {
            token.ThrowIfCancellationRequested();
            bool kept = !TargetsRemovedPage(objects, value, removedPages, token);
            survives[name] = survives.TryGetValue(name, out bool previous) ? previous || kept : kept;
        }
    }

    private static void ApplyNavigation(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog,
        PdfPageExtractor.CatalogRewriteState state, HashSet<int> removedPages, HashSet<string> removedNames,
        HashSet<int> removedAnnotations, ICollection<string>? report, CancellationToken token) {
        Set("PageLabels", state.PageLabels);
        // Reordering preserves destinations and arbitrary actions verbatim. Only removal needs pruning.
        if (removedPages.Count == 0) return;
        PruneNamedDestinations(objects, catalog, removedPages, report, token);

        // A destination is kept unless it provably goes to a removed page; unresolved and non-page destinations stay as found.
        bool Exists(PdfObject destination) {
            var seen = new HashSet<PdfObject>();
            while (seen.Add(destination)) {
                token.ThrowIfCancellationRequested();
                var resolved = PdfObjectLookup.Resolve(objects, destination);
                if (resolved is PdfDictionary dictionary && dictionary.Items.TryGetValue("D", out destination!)) continue;
                if (resolved is PdfArray array) return array.Items.Count == 0 || array.Items[0] is not PdfReference page || !removedPages.Contains(page.ObjectNumber);
                string? name = resolved is PdfName pdfName ? pdfName.Name : (resolved as PdfStringObj)?.Value;
                return name == null || !removedNames.Contains(name);
            }
            throw new InvalidOperationException("Page organization found a cyclic destination.");
        }

        var visitedOutlines = new HashSet<int>();
        var droppedOutlines = new HashSet<int>();
        if (catalog.Items.TryGetValue("Outlines", out var outlinesValue) && PdfObjectLookup.Resolve(objects, outlinesValue) is PdfDictionary outlines) {
            int count = PruneSiblings(outlines, outlinesValue);
            if (count == 0) outlines.Items.Remove("Count"); else outlines.Items["Count"] = new PdfNumber(count);
            if (!outlines.Items.ContainsKey("First")) { catalog.Items.Remove("Outlines");
                if (catalog.Get<PdfName>("PageMode")?.Name == "UseOutlines") catalog.Items.Remove("PageMode"); }
        }
        int PruneSiblings(PdfDictionary parent, PdfObject parentReference) {
            var retained = new List<(PdfReference Reference, PdfDictionary Dictionary, int Visible)>();
            PdfObject? next = parent.Items.TryGetValue("First", out var first) ? first : null;
            while (next != null) {
                token.ThrowIfCancellationRequested();
                if (next is not PdfReference reference || !visitedOutlines.Add(reference.ObjectNumber) ||
                    PdfObjectLookup.Resolve(objects, reference) is not PdfDictionary node)
                    throw new InvalidOperationException("Page organization requires an unambiguous outline tree.");
                next = node.Items.TryGetValue("Next", out var sibling) ? sibling : null;
                int children = PruneSiblings(node, reference);
                bool removed = node.Items.TryGetValue("Dest", out var destination) && !Exists(destination);
                if (removed) node.Items.Remove("Dest");
                if (node.Items.TryGetValue("A", out var action) && PdfObjectLookup.Resolve(objects, action) is PdfDictionary actionDictionary &&
                    actionDictionary.Get<PdfName>("S")?.Name == "GoTo" && actionDictionary.Items.TryGetValue("D", out var target) && !Exists(target)) {
                    if (actionDictionary.Items.ContainsKey("Next"))
                        throw new InvalidOperationException("A removed bookmark destination has a chained action that cannot be pruned safely.");
                    node.Items.Remove("A"); removed = true;
                }
                string title = node.Get<PdfStringObj>("Title")?.Value ?? string.Empty;
                if (removed && children == 0) {
                    droppedOutlines.Add(reference.ObjectNumber);
                    report?.Add(PdfPageNavigationReport.BookmarkRemoved(title));
                    continue;
                }
                if (removed) report?.Add(PdfPageNavigationReport.BookmarkDestinationRemoved(title));
                bool open = (node.Get<PdfNumber>("Count")?.Value ?? 0) >= 0;
                if (children == 0) node.Items.Remove("Count"); else node.Items["Count"] = new PdfNumber(open ? children : -children);
                retained.Add((reference, node, 1 + (open ? children : 0)));
            }
            parent.Items.Remove("First"); parent.Items.Remove("Last");
            for (int index = 0; index < retained.Count; index++) {
                var entry = retained[index]; entry.Dictionary.Items["Parent"] = parentReference;
                entry.Dictionary.Items.Remove("Prev"); entry.Dictionary.Items.Remove("Next");
                if (index > 0) entry.Dictionary.Items["Prev"] = retained[index - 1].Reference;
                if (index + 1 < retained.Count) entry.Dictionary.Items["Next"] = retained[index + 1].Reference;
            }
            if (retained.Count > 0) { parent.Items["First"] = retained[0].Reference; parent.Items["Last"] = retained[retained.Count - 1].Reference; }
            return retained.Sum(item => item.Visible);
        }

        var actionResults = new Dictionary<PdfDictionary, PdfObject?>();
        var activeActions = new HashSet<PdfDictionary>();
        PdfObject? PruneAction(PdfObject action) {
            if (PdfObjectLookup.Resolve(objects, action) is not PdfDictionary dictionary) return action;
            if (actionResults.TryGetValue(dictionary, out var previous)) return previous;
            if (!activeActions.Add(dictionary)) throw new InvalidOperationException("Page organization found a cyclic action chain.");
            token.ThrowIfCancellationRequested();
            if (dictionary.Items.TryGetValue("Next", out var next)) {
                if (PdfObjectLookup.Resolve(objects, next) is PdfArray array) {
                    var kept = new PdfArray(); foreach (var item in array.Items) { var value = PruneAction(item); if (value != null) kept.Items.Add(value); }
                    if (kept.Items.Count == 0) dictionary.Items.Remove("Next"); else dictionary.Items["Next"] = kept;
                } else { var value = PruneAction(next); if (value == null) dictionary.Items.Remove("Next"); else dictionary.Items["Next"] = value; }
            }
            if (dictionary.Get<PdfName>("S")?.Name != "GoTo" || !dictionary.Items.TryGetValue("D", out var destination) || Exists(destination)) {
                activeActions.Remove(dictionary); actionResults.Add(dictionary, action); return action;
            }
            if (dictionary.Items.ContainsKey("Next")) throw new InvalidOperationException("A removed page destination has a chained action that cannot be pruned safely.");
            activeActions.Remove(dictionary); actionResults.Add(dictionary, null);
            return null;
        }
        int removedActions = 0;
        foreach (var indirect in objects.Values) {
            token.ThrowIfCancellationRequested();
            if (indirect.Value is not PdfDictionary dictionary || removedPages.Contains(indirect.ObjectNumber)) continue;
            // Objects that leave the output with a removed page or bookmark are not reported as changed.
            bool reported = !removedAnnotations.Contains(indirect.ObjectNumber) && !droppedOutlines.Contains(indirect.ObjectNumber);
            foreach (string key in new[] { "A", "OpenAction" }) if (dictionary.Items.TryGetValue(key, out var action)) {
                var resolved = PdfObjectLookup.Resolve(objects, action);
                var value = resolved is PdfArray ? Exists(action) ? action : null : PruneAction(action);
                if (value == null) {
                    dictionary.Items.Remove(key);
                    if (!reported) continue;
                    if (ReferenceEquals(dictionary, catalog) && key == "OpenAction") report?.Add(PdfPageNavigationReport.OpenActionRemoved);
                    else removedActions++;
                } else dictionary.Items[key] = value;
            }
            if (dictionary.Items.TryGetValue("AA", out var additional) && PdfObjectLookup.Resolve(objects, additional) is PdfDictionary actions)
                foreach (string key in actions.Items.Keys.ToArray()) {
                    var value = PruneAction(actions.Items[key]);
                    if (value == null) { actions.Items.Remove(key); if (reported) removedActions++; } else actions.Items[key] = value;
                }
        }
        if (removedActions > 0) report?.Add(PdfPageNavigationReport.ActionsRemoved(removedActions));
        void Set(string key, PdfObject? value) { if (value is null) catalog.Items.Remove(key); else catalog.Items[key] = value; }
    }

    /// <summary>
    /// Removes named destinations that go to removed pages from the catalog's /Dests dictionary and /Names /Dests tree in
    /// place, keeping every other entry, the tree shape and its order. A tree that cannot be traversed is left as found;
    /// the surviving-reference check then refuses the output if it still reaches a removed page.
    /// </summary>
    private static void PruneNamedDestinations(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog,
        HashSet<int> removedPages, ICollection<string>? report, CancellationToken token) {
        int removed = 0;
        if (catalog.Items.TryGetValue("Dests", out var directValue) && PdfObjectLookup.Resolve(objects, directValue) is PdfDictionary direct) {
            foreach (string name in direct.Items.Keys.ToArray()) {
                token.ThrowIfCancellationRequested();
                if (!TargetsRemovedPage(objects, direct.Items[name], removedPages, token)) continue;
                direct.Items.Remove(name); removed++;
            }
            if (direct.Items.Count == 0) catalog.Items.Remove("Dests");
        }

        if (TryGetDestinationNameTree(objects, catalog, out var names, out var tree) &&
            TryCollectNameTreePairs(objects, tree, new List<(PdfStringObj, PdfObject)>(), token)) {
            if (PruneNameTreeNode(tree, isRoot: true).Remaining == 0) {
                names.Items.Remove("Dests");
                if (names.Items.Count == 0) catalog.Items.Remove("Names");
            }
        }
        if (removed > 0) report?.Add(PdfPageNavigationReport.NamedDestinationsRemoved(removed));

        (int Remaining, PdfStringObj? First, PdfStringObj? Last) PruneNameTreeNode(PdfDictionary node, bool isRoot) {
            token.ThrowIfCancellationRequested();
            int remaining = 0; PdfStringObj? first = null, last = null;
            if (node.Items.TryGetValue("Names", out var namesValue) && PdfObjectLookup.Resolve(objects, namesValue) is PdfArray pairs) {
                var kept = new PdfArray();
                for (int index = 0; index + 1 < pairs.Items.Count; index += 2) {
                    if (TargetsRemovedPage(objects, pairs.Items[index + 1], removedPages, token)) { removed++; continue; }
                    kept.Items.Add(pairs.Items[index]); kept.Items.Add(pairs.Items[index + 1]);
                    var key = (PdfStringObj)pairs.Items[index]; first ??= key; last = key; remaining++;
                }
                node.Items["Names"] = kept;
            }
            if (node.Items.TryGetValue("Kids", out var kidsValue) && PdfObjectLookup.Resolve(objects, kidsValue) is PdfArray kids) {
                var kept = new PdfArray();
                foreach (var kid in kids.Items) {
                    var child = PruneNameTreeNode((PdfDictionary)PdfObjectLookup.Resolve(objects, kid)!, isRoot: false);
                    if (child.Remaining == 0) continue;
                    kept.Items.Add(kid); first ??= child.First; last = child.Last; remaining += child.Remaining;
                }
                node.Items["Kids"] = kept;
            }
            if (!isRoot && remaining > 0 && node.Items.ContainsKey("Limits")) {
                var limits = new PdfArray(); limits.Items.Add(first!); limits.Items.Add(last!); node.Items["Limits"] = limits;
            }
            return (remaining, first, last);
        }
    }

    private static bool TryGetDestinationNameTree(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog,
        out PdfDictionary names, out PdfDictionary tree) {
        names = null!; tree = null!;
        if (!catalog.Items.TryGetValue("Names", out var namesValue) || PdfObjectLookup.Resolve(objects, namesValue) is not PdfDictionary namesDictionary ||
            !namesDictionary.Items.TryGetValue("Dests", out var treeValue) || PdfObjectLookup.Resolve(objects, treeValue) is not PdfDictionary treeDictionary)
            return false;
        names = namesDictionary; tree = treeDictionary;
        return true;
    }

    /// <summary>Collects name-tree pairs, returning false for a tree whose shape, depth or size cannot be traversed safely.</summary>
    private static bool TryCollectNameTreePairs(Dictionary<int, PdfIndirectObject> objects, PdfDictionary root,
        List<(PdfStringObj Name, PdfObject Value)> pairs, CancellationToken token) {
        var visited = new HashSet<PdfDictionary>();
        int nodes = 0;
        return Visit(root, 0);

        bool Visit(PdfDictionary node, int depth) {
            token.ThrowIfCancellationRequested();
            if (depth > PdfReadLimits.DefaultMaxNameTreeDepth || ++nodes > PdfReadLimits.DefaultMaxNameTreeNodes || !visited.Add(node)) return false;
            bool hasNames = node.Items.TryGetValue("Names", out var namesValue);
            bool hasKids = node.Items.TryGetValue("Kids", out var kidsValue);
            if (hasNames) {
                if (PdfObjectLookup.Resolve(objects, namesValue) is not PdfArray array || array.Items.Count % 2 != 0) return false;
                for (int index = 0; index < array.Items.Count; index += 2) {
                    if (array.Items[index] is not PdfStringObj name) return false;
                    pairs.Add((name, array.Items[index + 1]));
                }
            }
            if (hasKids) {
                if (PdfObjectLookup.Resolve(objects, kidsValue) is not PdfArray kids) return false;
                foreach (var kid in kids.Items)
                    if (PdfObjectLookup.Resolve(objects, kid) is not PdfDictionary child || !Visit(child, depth + 1)) return false;
            }
            return hasNames || hasKids;
        }
    }

    private static bool TargetsRemovedPage(Dictionary<int, PdfIndirectObject> objects, PdfObject destination, HashSet<int> removedPages,
        CancellationToken token) {
        var seen = new HashSet<PdfObject>();
        PdfObject? current = destination;
        while (current is not null && seen.Add(current)) {
            token.ThrowIfCancellationRequested();
            var resolved = PdfObjectLookup.Resolve(objects, current);
            if (resolved is PdfDictionary dictionary) { current = dictionary.Items.TryGetValue("D", out var inner) ? inner : null; continue; }
            return resolved is PdfArray array && array.Items.Count > 0 && array.Items[0] is PdfReference page && removedPages.Contains(page.ObjectNumber);
        }
        return false;
    }
}

using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfTaggedPageOrganizer {
    private static void ApplyNavigation(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog,
        PdfPageExtractor.CatalogRewriteState state, HashSet<int> removedPages, CancellationToken token) {
        Set("PageLabels", state.PageLabels);
        // Reordering preserves destinations and arbitrary actions verbatim. Only removal needs pruning.
        if (removedPages.Count == 0) return;
        Set("Dests", state.NamedDestinations);
        if (catalog.Items.TryGetValue("Names", out var namesValue) && PdfObjectLookup.Resolve(objects, namesValue) is PdfDictionary names) {
            if (state.NamedDestinationNameTree is null) names.Items.Remove("Dests");
            else names.Items["Dests"] = state.NamedDestinationNameTree;
        }
        var named = new Dictionary<string, PdfObject>(StringComparer.Ordinal);
        if (state.NamedDestinations is PdfDictionary direct)
            foreach (var pair in direct.Items) named[pair.Key] = pair.Value;
        if (state.NamedDestinationNameTree is PdfDictionary tree && tree.Get<PdfArray>("Names") is PdfArray pairs)
            for (int index = 0; index + 1 < pairs.Items.Count; index += 2)
                if (pairs.Items[index] is PdfStringObj name) named[name.Value] = pairs.Items[index + 1];

        bool Exists(PdfObject destination) {
            var seen = new HashSet<PdfObject>();
            while (seen.Add(destination)) {
                token.ThrowIfCancellationRequested();
                var resolved = PdfObjectLookup.Resolve(objects, destination);
                if (resolved is PdfDictionary dictionary && dictionary.Items.TryGetValue("D", out destination!)) continue;
                if (resolved is PdfArray array && array.Items.Count > 0 && array.Items[0] is PdfReference page)
                    return !removedPages.Contains(page.ObjectNumber);
                string? name = resolved is PdfName pdfName ? pdfName.Name : (resolved as PdfStringObj)?.Value;
                if (name == null || !named.TryGetValue(name, out destination!)) return false;
            }
            throw new InvalidOperationException("Page organization found a cyclic destination.");
        }

        var visitedOutlines = new HashSet<int>();
        if (catalog.Items.TryGetValue("Outlines", out var outlinesValue) && PdfObjectLookup.Resolve(objects, outlinesValue) is PdfDictionary outlines) {
            int count = PruneSiblings(outlines, outlinesValue);
            outlines.Items["Count"] = new PdfNumber(count);
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
                if (removed && children == 0) continue;
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
        foreach (var indirect in objects.Values) {
            token.ThrowIfCancellationRequested();
            if (indirect.Value is not PdfDictionary dictionary || removedPages.Contains(indirect.ObjectNumber)) continue;
            foreach (string key in new[] { "A", "OpenAction" }) if (dictionary.Items.TryGetValue(key, out var action)) {
                var resolved = PdfObjectLookup.Resolve(objects, action);
                var value = resolved is PdfArray ? Exists(action) ? action : null : PruneAction(action);
                if (value == null) dictionary.Items.Remove(key); else dictionary.Items[key] = value;
            }
            if (dictionary.Items.TryGetValue("AA", out var additional) && PdfObjectLookup.Resolve(objects, additional) is PdfDictionary actions)
                foreach (string key in actions.Items.Keys.ToArray()) { var value = PruneAction(actions.Items[key]); if (value == null) actions.Items.Remove(key); else actions.Items[key] = value; }
        }
        void Set(string key, PdfObject? value) { if (value is null) catalog.Items.Remove(key); else catalog.Items[key] = value; }
    }
}

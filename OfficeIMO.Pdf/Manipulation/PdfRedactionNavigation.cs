using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Proves the navigation and action subset retained by redaction, independently of raw marker spellings.</summary>
internal static partial class PdfRedactionNavigation {
    internal static bool CanPreserve(PdfReadDocument document, CancellationToken cancellationToken) =>
        Inspect(document, navigationOnly: false, cancellationToken);

    /// <summary>Plans navigation without changing the existing form/script planning contract; mutation still checks the whole graph.</summary>
    internal static bool CanPlanNavigation(PdfReadDocument document, CancellationToken cancellationToken) =>
        Inspect(document, navigationOnly: true, cancellationToken);

    private static bool Inspect(PdfReadDocument document, bool navigationOnly, CancellationToken cancellationToken) {
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        var outlines = new HashSet<PdfDictionary>(OutlineItems(objects, document.CatalogDictionary, cancellationToken));
        // URI link annotations already have search/removal handling. They are not opening or outline actions.
        var linkActions = new HashSet<PdfObject>();
        foreach (PdfIndirectObject item in objects.Values) {
            if (item.Value is PdfDictionary annotation && LinkOnly(annotation) &&
                annotation.Items.TryGetValue("A", out PdfObject? link) && PdfObjectLookup.Resolve(objects, link) is PdfDictionary action) linkActions.Add(action);
        }
        foreach (PdfReadPage page in document.Pages) {
            if (!objects.TryGetValue(page.ObjectNumber, out PdfIndirectObject? item) || item.Value is not PdfDictionary dictionary ||
                !dictionary.Items.TryGetValue("Annots", out PdfObject? annotations) || PdfObjectLookup.Resolve(objects, annotations) is not PdfArray array) continue;
            foreach (PdfObject value in array.Items) {
                if (PdfObjectLookup.Resolve(objects, value) is PdfDictionary annotation && LinkOnly(annotation) &&
                    annotation.Items.TryGetValue("A", out PdfObject? link) && PdfObjectLookup.Resolve(objects, link) is PdfDictionary action) linkActions.Add(action);
            }
        }
        var completed = new HashSet<PdfObject>();
        var active = new HashSet<PdfObject>();
        var visited = new HashSet<PdfObject>();
        var pending = new Stack<PdfObject>(objects.Values.Select(static item => item.Value));
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject value = pending.Pop();
            if (!visited.Add(value)) continue;
            if (!navigationOnly && value is PdfName name && PdfActiveContentPolicy.MarkerNames.Contains(name.Name) && name.Name != "AA") return false;
            if (value is PdfArray array) { foreach (PdfObject item in array.Items) pending.Push(item); continue; }
            if (value is PdfStream stream) { pending.Push(stream.Dictionary); continue; }
            if (value is not PdfDictionary dictionary) continue;
            // Presentation nodes introduce separate executable /NA and /PA roots. This class is not qualified.
            if (dictionary.Items.ContainsKey("PresSteps") || dictionary.Get<PdfName>("Type")?.Name == "NavNode") return false;
            if (!navigationOnly && dictionary.Items.Keys.Any(key => key != "AA" && PdfActiveContentPolicy.MarkerNames.Contains(key))) return false;
            if (dictionary.Items.TryGetValue("OpenAction", out PdfObject? opening) &&
                !(Resolve(opening) is PdfArray destination ? Destination(destination) : Action(opening, 0))) return false;
            if ((!navigationOnly || NavigationOwner(dictionary)) && dictionary.Items.TryGetValue("AA", out PdfObject? additional)) {
                if (Resolve(additional) is not PdfDictionary actions || actions.Items.Values.Any(item => !Action(item, 0))) return false;
            }
            if ((!navigationOnly || outlines.Contains(dictionary)) && dictionary.Items.TryGetValue("A", out PdfObject? action) &&
                (outlines.Contains(dictionary) || dictionary.Items.ContainsKey("Subtype") || dictionary.Items.ContainsKey("FT") ||
                 Resolve(action) is PdfDictionary candidate && candidate.Items.ContainsKey("S")) && !(LinkOnly(dictionary) ? LinkAction(action) : Action(action, 0))) return false;
            if (!navigationOnly && dictionary.Get<PdfName>("Type")?.Name == "Action" && !(linkActions.Contains(dictionary) && !outlines.Contains(dictionary) && !NavigationOwner(dictionary) && !dictionary.Items.ContainsKey("FT") ? LinkAction(dictionary) : Action(dictionary, 0))) return false;
            // All indirect objects are inspected separately. Do not follow page/tree cycles or compressed bytes here.
            foreach (PdfObject item in dictionary.Items.Values) pending.Push(item);
        }
        return true;

        bool NavigationOwner(PdfDictionary dictionary) => ReferenceEquals(dictionary, document.CatalogDictionary) || dictionary.Get<PdfName>("Type")?.Name is "Catalog" or "Page";
        bool LinkOnly(PdfDictionary dictionary) => dictionary.Get<PdfName>("Subtype")?.Name == "Link" &&
            !outlines.Contains(dictionary) && !NavigationOwner(dictionary) && !dictionary.Items.ContainsKey("FT");
        PdfObject? Resolve(PdfObject? value) => PdfObjectLookup.Resolve(objects, value);
        bool LinkAction(PdfObject? value) {
            if (Resolve(value) is not PdfDictionary action || Resolve(action.Items.TryGetValue("S", out PdfObject? type) ? type : null) is not PdfName { Name: "URI" }) return Action(value, 0);
            if (!action.Items.TryGetValue("URI", out PdfObject? uri) || Resolve(uri) is not PdfStringObj) return false;
            return !action.Items.TryGetValue("Next", out PdfObject? next) ||
                (Resolve(next) is PdfArray sequence ? sequence.Items.All(item => Action(item, 0)) : Action(next, 0));
        }

        bool Action(PdfObject? value, int depth) {
            cancellationToken.ThrowIfCancellationRequested();
            if (depth > document.ReadOptions.Limits.MaxObjectNestingDepth) return false;
            value = Resolve(value);
            if (value is null || !active.Add(value)) return false;
            if (completed.Contains(value)) { active.Remove(value); return true; }
            bool safe = value is PdfDictionary action &&
                Resolve(action.Items.TryGetValue("S", out PdfObject? type) ? type : null) is PdfName { Name: "GoTo" } &&
                action.Items.TryGetValue("D", out PdfObject? target) && Destination(target);
            if (safe && ((PdfDictionary)value!).Items.TryGetValue("Next", out PdfObject? next)) {
                safe = Resolve(next) is PdfArray sequence ? sequence.Items.All(item => Action(item, depth + 1)) : Action(next, depth + 1);
            }
            active.Remove(value);
            if (safe) completed.Add(value);
            return safe;
        }
        bool Destination(PdfObject? value) {
            value = Resolve(value);
            if (value is PdfStringObj or PdfName) {
                string name = value is PdfStringObj text ? text.Value : ((PdfName)value).Name;
                return document.UncheckedNamedDestinations.Any(destination => destination.Name == name && destination.PageNumber.HasValue);
            }
            if (value is not PdfArray array || array.Items.Count < 2 || array.Items[0] is not PdfReference page ||
                !PdfObjectLookup.TryGet(objects, page, out _) || !document.Pages.Any(candidate => candidate.ObjectNumber == page.ObjectNumber) ||
                Resolve(array.Items[1]) is not PdfName mode) return false;
            int count = mode.Name switch { "XYZ" => 5, "Fit" or "FitB" => 2, "FitH" or "FitV" or "FitBH" or "FitBV" => 3, "FitR" => 6, _ => -1 };
            return array.Items.Count == count && array.Items.Skip(2).All(item => Resolve(item) is PdfNumber or PdfNull);
        }
    }

    /// <summary>Traverses the live outline links without relying on the reading projection's item limit.</summary>
    internal static IEnumerable<PdfDictionary> OutlineItems(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? catalog, CancellationToken cancellationToken) {
        if (catalog is null || !catalog.Items.TryGetValue("Outlines", out PdfObject? value) ||
            PdfObjectLookup.Resolve(objects, value) is not PdfDictionary root || !root.Items.TryGetValue("First", out PdfObject? first)) yield break;
        var pending = new Stack<PdfObject>(); pending.Push(first);
        var visited = new HashSet<PdfDictionary>();
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            if (PdfObjectLookup.Resolve(objects, pending.Pop()) is not PdfDictionary item || !visited.Add(item)) continue;
            yield return item;
            if (item.Items.TryGetValue("Next", out PdfObject? next)) pending.Push(next);
            if (item.Items.TryGetValue("First", out PdfObject? child)) pending.Push(child);
        }
    }
}

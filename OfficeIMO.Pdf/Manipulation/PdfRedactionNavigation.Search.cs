using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionNavigation {
    /// <summary>Finds titles to clear and identifiers whose renaming would break destination bindings.</summary>
    internal static (List<PdfDictionary> Titles, bool DestinationMatch) Search(Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary? catalog, PdfRedactionSearchCriteria[] criteria, CancellationToken cancellationToken) {
        var titles = new List<PdfDictionary>();
        if (criteria.Length == 0 || catalog is null) return (titles, false);
        bool destinationMatch = false;
        foreach (PdfDictionary item in OutlineItems(objects, catalog, cancellationToken)) {
            if (item.Items.TryGetValue("Title", out PdfObject? title) && Holds(title)) titles.Add(item);
            if (item.Items.TryGetValue("Dest", out PdfObject? destination)) destinationMatch |= Holds(destination);
        }
        foreach (PdfIndirectObject item in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (item.Value is PdfDictionary dictionary) {
                if (dictionary.Get<PdfName>("S")?.Name == "GoTo" && dictionary.Items.TryGetValue("D", out PdfObject? target)) destinationMatch |= Holds(target);
                if (dictionary.Items.TryGetValue("OpenAction", out PdfObject? opening)) destinationMatch |= Holds(opening);
            }
        }
        if (catalog.Items.TryGetValue("Dests", out PdfObject? legacy) && Resolve(legacy) is PdfDictionary destinations)
            destinationMatch |= destinations.Items.Any(pair => criteria.Any(criterion => criterion.Matches(pair.Key)) || Holds(pair.Value));
        if (catalog.Items.TryGetValue("Names", out PdfObject? names) && Resolve(names) is PdfDictionary trees && trees.Items.TryGetValue("Dests", out PdfObject? tree)) {
            var pending = new Stack<PdfObject>(); pending.Push(tree);
            var visited = new HashSet<PdfDictionary>();
            while (pending.Count > 0) {
                cancellationToken.ThrowIfCancellationRequested();
                if (Resolve(pending.Pop()) is not PdfDictionary node || !visited.Add(node)) continue;
                if (node.Items.TryGetValue("Names", out PdfObject? entries) && Resolve(entries) is PdfArray array) destinationMatch |= array.Items.Any(Holds);
                if (node.Items.TryGetValue("Limits", out PdfObject? limits) && Resolve(limits) is PdfArray bounds) destinationMatch |= bounds.Items.Any(Holds);
                if (node.Items.TryGetValue("Kids", out PdfObject? children) && Resolve(children) is PdfArray kids)
                    foreach (PdfObject child in kids.Items) pending.Push(child);
            }
        }
        return (titles, destinationMatch);

        PdfObject? Resolve(PdfObject? value) => PdfObjectLookup.Resolve(objects, value);
        bool Holds(PdfObject? value) {
            value = Resolve(value);
            // A destination dictionary wraps one /D; do not follow arbitrary recursive dictionaries.
            if (value is PdfDictionary dictionary && dictionary.Items.TryGetValue("D", out PdfObject? target)) value = Resolve(target);
            return value switch {
                PdfStringObj text => criteria.Any(criterion => criterion.Matches(text.Value)),
                PdfName name => criteria.Any(criterion => criterion.Matches(name.Name)),
                _ => false
            };
        }
    }
}

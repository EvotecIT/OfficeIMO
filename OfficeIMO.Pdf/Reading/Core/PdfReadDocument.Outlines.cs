namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private IReadOnlyList<PdfOutlineItem> ExtractOutlines() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("Outlines", out var outlinesObj) ||
            ResolveDict(outlinesObj) is not PdfDictionary outlines ||
            !outlines.Items.TryGetValue("First", out var firstObj)) {
            return Array.Empty<PdfOutlineItem>();
        }

        var visited = new HashSet<int>();
        return ReadOutlineSiblings(firstObj, 1, visited).AsReadOnly();
    }

    private List<PdfOutlineItem> ReadOutlineSiblings(PdfObject firstObj, int level, HashSet<int> visited) {
        var items = new List<PdfOutlineItem>();
        PdfObject? currentObj = firstObj;

        while (currentObj is not null && ResolveDict(currentObj) is PdfDictionary current) {
            int objectNumber = currentObj is PdfReference reference ? reference.ObjectNumber : FindObjectNumberFor(current);
            if (objectNumber > 0 && !visited.Add(objectNumber)) {
                break;
            }

            string title = current.Get<PdfStringObj>("Title")?.Value ?? string.Empty;
            var (pageNumber, destinationTop, destinationMode, destinationLeft, destinationBottom, destinationRight) = GetOutlineDestination(current);
            bool isExpanded = !current.Items.TryGetValue("Count", out var countObject) ||
                ResolveObject(countObject) is not PdfNumber countNumber ||
                countNumber.Value >= 0D;
            var children = current.Items.TryGetValue("First", out var childObj)
                ? ReadOutlineSiblings(childObj, level + 1, visited)
                : new List<PdfOutlineItem>();

            items.Add(new PdfOutlineItem(title, level, pageNumber, destinationTop, isExpanded, children.AsReadOnly(), destinationMode, destinationLeft, destinationBottom, destinationRight));

            currentObj = current.Items.TryGetValue("Next", out var nextObj) ? nextObj : null;
        }

        return items;
    }
}

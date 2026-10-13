using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    private static void NormalizeStructureBounds(Dictionary<int, PdfIndirectObject> objects, PdfObject structure,
        IReadOnlyDictionary<int, PageResizeTransform> transforms, CancellationToken cancellationToken) {
        var root = PdfObjectLookup.Resolve(objects, structure) as PdfDictionary;
        PdfDictionary? classes = root is not null && root.Items.TryGetValue("ClassMap", out PdfObject? classValue)
            ? PdfObjectLookup.Resolve(objects, classValue) as PdfDictionary : null;
        var mappedClasses = new Dictionary<(string Name, int Page), string>();
        var visited = new HashSet<PdfObject>();
        var pending = new Stack<(PdfObject Value, int? Page)>();
        pending.Push((structure, null));
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var entry = pending.Pop();
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, entry.Value);
            if (resolved is null || !visited.Add(resolved)) continue;
            if (resolved is PdfArray array) {
                foreach (PdfObject child in array.Items) pending.Push((child, entry.Page));
                continue;
            }
            if (resolved is not PdfDictionary dictionary) continue;
            int? page = dictionary.Get<PdfReference>("Pg")?.ObjectNumber ?? entry.Page;
            bool hasBounds = dictionary.Items.TryGetValue("A", out PdfObject? attributes) && HasStructureBounds(objects, attributes);
            bool hasClassBounds = classes is not null && dictionary.Items.TryGetValue("C", out PdfObject? names) &&
                HasStructureClassBounds(objects, names, classes);
            if (hasBounds || hasClassBounds) page ??= FindSingleStructurePage(objects, dictionary, cancellationToken);
            if (page.HasValue && transforms.TryGetValue(page.Value, out PageResizeTransform transform) &&
                dictionary.Items.TryGetValue("A", out attributes))
                dictionary.Items["A"] = NormalizeStructureAttributes(objects, attributes, transform);
            if (page.HasValue && transforms.TryGetValue(page.Value, out transform) && classes is not null &&
                dictionary.Items.TryGetValue("C", out PdfObject? classNames))
                dictionary.Items["C"] = NormalizeStructureClasses(objects, classNames, classes, mappedClasses, transform);
            if (!page.HasValue && (hasBounds || hasClassBounds))
                throw new NotSupportedException("Rotation normalization requires a page association for structure layout bounds.");
            if (dictionary.Items.TryGetValue("K", out PdfObject? kids)) pending.Push((kids, page));
        }
    }

    private static bool HasStructureClassBounds(Dictionary<int, PdfIndirectObject> objects, PdfObject value, PdfDictionary classes, int depth = 0) {
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
        if (resolved is PdfArray && depth > 0) throw new InvalidOperationException("Structure class names require a flat array.");
        return resolved is PdfArray array ? array.Items.Any(item => HasStructureClassBounds(objects, item, classes, depth + 1)) :
            resolved is PdfName name && classes.Items.TryGetValue(name.Name, out PdfObject? attributes) && HasStructureBounds(objects, attributes);
    }

    private static PdfObject NormalizeStructureClasses(Dictionary<int, PdfIndirectObject> objects, PdfObject value,
        PdfDictionary classes, Dictionary<(string Name, int Page), string> mapped, PageResizeTransform transform, int depth = 0) {
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
        if (resolved is PdfArray && depth > 0) throw new InvalidOperationException("Structure class names require a flat array.");
        if (resolved is PdfArray array) {
            var copy = new PdfArray();
            foreach (PdfObject item in array.Items) copy.Items.Add(NormalizeStructureClasses(objects, item, classes, mapped, transform, depth + 1));
            return copy;
        }
        if (resolved is not PdfName name || !classes.Items.TryGetValue(name.Name, out PdfObject? attributes) ||
            !HasStructureBounds(objects, attributes)) return value;
        var key = (name.Name, transform.PageObjectNumber);
        if (!mapped.TryGetValue(key, out string? replacement)) {
            int index = classes.Items.Count;
            do { replacement = "OfficeIMORotation" + (++index).ToString(System.Globalization.CultureInfo.InvariantCulture); }
            while (classes.Items.ContainsKey(replacement));
            classes.Items[replacement] = NormalizeStructureAttributes(objects, attributes, transform);
            mapped.Add(key, replacement);
        }
        return new PdfName(replacement);
    }

    private static bool HasStructureBounds(Dictionary<int, PdfIndirectObject> objects, PdfObject value, int depth = 0) {
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
        if (resolved is PdfArray && depth > 0) throw new InvalidOperationException("Structure attributes require a flat array.");
        return resolved is PdfArray array ? array.Items.Any(item => HasStructureBounds(objects, item, depth + 1)) :
            resolved is PdfDictionary dictionary && dictionary.Get<PdfName>("O")?.Name == "Layout" && dictionary.Items.ContainsKey("BBox");
    }

    private static int? FindSingleStructurePage(Dictionary<int, PdfIndirectObject> objects, PdfDictionary element,
        CancellationToken cancellationToken) {
        if (!element.Items.TryGetValue("K", out PdfObject? kids)) return null;
        var visited = new HashSet<PdfObject>();
        var pending = new Stack<PdfObject>(); pending.Push(kids);
        int? page = null;
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject? current = PdfObjectLookup.Resolve(objects, pending.Pop());
            if (current is null || !visited.Add(current)) continue;
            if (current is PdfArray array) { foreach (PdfObject child in array.Items) pending.Push(child); }
            else if (current is PdfDictionary dictionary) {
                if (dictionary.Get<PdfReference>("Pg") is PdfReference reference) {
                    if (page.HasValue && page != reference.ObjectNumber) return null;
                    page = reference.ObjectNumber;
                }
                if (dictionary.Items.TryGetValue("K", out PdfObject? children)) pending.Push(children);
            }
        }
        return page;
    }

    private static PdfObject NormalizeStructureAttributes(Dictionary<int, PdfIndirectObject> objects,
        PdfObject attributes, PageResizeTransform transform, int depth = 0) {
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, attributes);
        if (resolved is PdfArray && depth > 0) throw new InvalidOperationException("Structure attributes require a flat array.");
        if (resolved is PdfArray array) {
            var result = new PdfArray();
            foreach (PdfObject item in array.Items) result.Items.Add(NormalizeStructureAttributes(objects, item, transform, depth + 1));
            return result;
        }
        if (resolved is not PdfDictionary dictionary) return attributes;
        var clone = (PdfDictionary)ClonePdfObject(dictionary);
        if (clone.Get<PdfName>("O")?.Name == "Layout" && clone.Items.TryGetValue("BBox", out PdfObject? bounds)) {
            double[] box = ReadNormalizationBox(objects, bounds);
            clone.Items["BBox"] = TransformNormalizationBox(box[0], box[1], box[2], box[3], ToMatrix(transform));
        }
        return clone;
    }
}

namespace OfficeIMO.Pdf;

internal static class PdfAssociatedFileGraph {
    internal static IReadOnlyList<PdfArray> FindAssociatedFileArrays(Dictionary<int, PdfIndirectObject> objects) {
        var arrays = new List<PdfArray>();
        var visited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject item in objects.Values) {
            CollectAssociatedFileArrays(objects, item.Value, arrays, visited);
        }

        return arrays;
    }

    internal static bool RemoveAssociatedFileReferences(Dictionary<int, PdfIndirectObject> objects) {
        bool changed = false;
        var visited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject item in objects.Values) {
            changed = RemoveAssociatedFileReferences(item.Value, visited) || changed;
        }

        return changed;
    }

    private static void CollectAssociatedFileArrays(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        List<PdfArray> arrays,
        HashSet<PdfObject> visited) {
        if (!visited.Add(value)) return;
        PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
        if (dictionary is not null) {
            if (dictionary.Items.TryGetValue("AF", out PdfObject? associatedFilesObject) &&
                PdfObjectLookup.Resolve(objects, associatedFilesObject) is PdfArray associatedFiles &&
                !arrays.Contains(associatedFiles)) {
                arrays.Add(associatedFiles);
            }

            foreach (PdfObject child in dictionary.Items.Values) {
                if (child is not PdfReference) CollectAssociatedFileArrays(objects, child, arrays, visited);
            }

            return;
        }

        if (value is PdfArray array) {
            foreach (PdfObject child in array.Items) {
                if (child is not PdfReference) CollectAssociatedFileArrays(objects, child, arrays, visited);
            }
        }
    }

    private static bool RemoveAssociatedFileReferences(PdfObject value, HashSet<PdfObject> visited) {
        if (!visited.Add(value)) return false;
        PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
        if (dictionary is not null) {
            bool changed = dictionary.Items.Remove("AF");
            foreach (PdfObject child in dictionary.Items.Values.ToArray()) {
                if (child is not PdfReference) changed = RemoveAssociatedFileReferences(child, visited) || changed;
            }

            return changed;
        }

        if (value is not PdfArray array) return false;
        bool arrayChanged = false;
        foreach (PdfObject child in array.Items) {
            if (child is not PdfReference) arrayChanged = RemoveAssociatedFileReferences(child, visited) || arrayChanged;
        }

        return arrayChanged;
    }
}

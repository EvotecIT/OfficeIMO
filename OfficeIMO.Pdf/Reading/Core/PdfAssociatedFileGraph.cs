namespace OfficeIMO.Pdf;

internal static class PdfAssociatedFileGraph {
    internal static bool RemoveAssociatedFileReferences(Dictionary<int, PdfIndirectObject> objects) {
        bool changed = false;
        var visited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject item in objects.Values) {
            changed = RemoveAssociatedFileReferences(item.Value, visited) || changed;
        }

        return changed;
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

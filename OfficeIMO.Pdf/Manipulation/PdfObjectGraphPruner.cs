namespace OfficeIMO.Pdf;

internal static class PdfObjectGraphPruner {
    public static void PruneUnreachableObjects(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber) {
        if (!objects.TryGetValue(catalogObjectNumber, out PdfIndirectObject? catalogObject)) {
            return;
        }

        var reachable = new HashSet<int>();
        CollectReachableObjectNumbers(objects, new PdfReference(catalogObjectNumber, catalogObject.Generation), reachable);
        foreach (int objectNumber in objects.Keys.ToArray()) {
            if (!reachable.Contains(objectNumber)) {
                objects.Remove(objectNumber);
            }
        }
    }

    private static void CollectReachableObjectNumbers(Dictionary<int, PdfIndirectObject> objects, PdfObject value, HashSet<int> reachable) {
        var pending = new Stack<PdfObject>();
        pending.Push(value);
        while (pending.Count > 0) {
            PdfObject current = pending.Pop();
            if (current is PdfReference reference) {
                if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
                    reachable.Add(indirect.ObjectNumber)) {
                    pending.Push(indirect.Value);
                }
                continue;
            }

            if (current is PdfArray array) {
                for (int i = array.Items.Count - 1; i >= 0; i--) pending.Push(array.Items[i]);
                continue;
            }

            if (current is PdfDictionary dictionary) {
                foreach (PdfObject child in dictionary.Items.Values) pending.Push(child);
                continue;
            }

            if (current is PdfStream stream) pending.Push(stream.Dictionary);
        }
    }
}

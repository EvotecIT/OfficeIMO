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

    private static void CollectReachableObjectNumbers(Dictionary<int, PdfIndirectObject> objects, PdfObject? value, HashSet<int> reachable) {
        if (value is PdfReference reference) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                !reachable.Add(indirect.ObjectNumber)) {
                return;
            }

            CollectReachableObjectNumbers(objects, indirect.Value, reachable);
            return;
        }

        if (value is PdfArray array) {
            for (int i = 0; i < array.Items.Count; i++) {
                CollectReachableObjectNumbers(objects, array.Items[i], reachable);
            }

            return;
        }

        if (value is PdfDictionary dictionary) {
            foreach (PdfObject child in dictionary.Items.Values) {
                CollectReachableObjectNumbers(objects, child, reachable);
            }

            return;
        }

        if (value is PdfStream stream) {
            CollectReachableObjectNumbers(objects, stream.Dictionary, reachable);
        }
    }
}

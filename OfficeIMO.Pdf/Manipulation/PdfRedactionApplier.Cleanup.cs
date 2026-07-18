namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private static bool ApplyCleanupPolicy(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber, PdfRedactionCleanupScope scope) {
        if (scope == PdfRedactionCleanupScope.None || !objects.TryGetValue(catalogObjectNumber, out PdfIndirectObject? catalogObject) || catalogObject.Value is not PdfDictionary catalog) return false;
        bool changed = false;
        if ((scope & PdfRedactionCleanupScope.Metadata) != 0) { changed = catalog.Items.Remove("Metadata") || changed; foreach (PdfIndirectObject item in objects.Values) if (item.Value is PdfDictionary dictionary && string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal)) changed = dictionary.Items.Remove("Metadata") || changed; }
        if ((scope & PdfRedactionCleanupScope.Attachments) != 0) changed = RemoveAttachments(objects, catalog) || changed;
        if ((scope & PdfRedactionCleanupScope.StructureTree) != 0) { changed = catalog.Items.Remove("StructTreeRoot") || changed; changed = catalog.Items.Remove("MarkInfo") || changed; foreach (PdfIndirectObject item in objects.Values) if (item.Value is PdfDictionary dictionary && string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal)) changed = dictionary.Items.Remove("StructParents") || changed; }
        if ((scope & PdfRedactionCleanupScope.AlternateText) != 0) foreach (PdfIndirectObject item in objects.Values) changed = RemoveKeys(item.Value, "Alt", "ActualText", "TU") || changed;
        if ((scope & PdfRedactionCleanupScope.OptionalContent) != 0) { changed = catalog.Items.Remove("OCProperties") || changed; foreach (PdfIndirectObject item in objects.Values) changed = RemoveOptionalContentReferences(item.Value) || changed; }
        return changed;
    }

    private static bool RemoveAttachments(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog) {
        bool changed = PdfAssociatedFileGraph.RemoveAssociatedFileReferences(objects);
        PdfDictionary? names = ResolveDictionary(objects, catalog.Items.TryGetValue("Names", out PdfObject? namesObject) ? namesObject : null);
        if (names is not null) { changed = names.Items.Remove("EmbeddedFiles") || changed; if (names.Items.Count == 0) changed = catalog.Items.Remove("Names") || changed; }
        foreach (PdfIndirectObject item in objects.Values) {
            if (item.Value is not PdfDictionary page || !string.Equals(page.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal) || !page.Items.TryGetValue("Annots", out PdfObject? annotsObject) || PdfObjectLookup.Resolve(objects, annotsObject) is not PdfArray annots) continue;
            for (int i = annots.Items.Count - 1; i >= 0; i--) { PdfDictionary? annotation = PdfObjectLookup.Resolve(objects, annots.Items[i]) as PdfDictionary; if (annotation is null || !string.Equals(annotation.Get<PdfName>("Subtype")?.Name, "FileAttachment", StringComparison.Ordinal)) continue; if (annots.Items[i] is PdfReference reference) objects.Remove(reference.ObjectNumber); annots.Items.RemoveAt(i); changed = true; }
            if (annots.Items.Count == 0) page.Items.Remove("Annots");
        }
        return changed;
    }

    private static bool RemoveOptionalContentReferences(PdfObject value) {
        PdfDictionary? dictionary = value is PdfDictionary direct ? direct : value is PdfStream stream ? stream.Dictionary : null; if (dictionary is null) return false;
        bool changed = dictionary.Items.Remove("OC");
        if (dictionary.Items.TryGetValue("Properties", out PdfObject? properties) && properties is PdfDictionary) changed = dictionary.Items.Remove("Properties") || changed;
        foreach (PdfObject child in dictionary.Items.Values.ToArray()) changed = RemoveOptionalContentReferences(child) || changed;
        return changed;
    }

    private static bool RemoveKeys(PdfObject value, params string[] keys) {
        bool changed = false; PdfDictionary? dictionary = value is PdfDictionary direct ? direct : value is PdfStream stream ? stream.Dictionary : null; if (dictionary is null) return false;
        for (int i = 0; i < keys.Length; i++) changed = dictionary.Items.Remove(keys[i]) || changed;
        foreach (PdfObject child in dictionary.Items.Values.ToArray()) changed = RemoveKeys(child, keys) || changed;
        return changed;
    }
}

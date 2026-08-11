namespace OfficeIMO.Pdf;

/// <summary>Owns optional image-mask dictionary semantics shared by image normalization paths.</summary>
internal static class PdfImageMaskSemantics {
    internal static bool HasSoftMask(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!dictionary.Items.TryGetValue("SMask", out PdfObject? softMaskObject)) return false;
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, softMaskObject);
        return resolved is not PdfNull &&
            (resolved is not PdfName softMaskName || !string.Equals(softMaskName.Name, "None", StringComparison.Ordinal));
    }
}

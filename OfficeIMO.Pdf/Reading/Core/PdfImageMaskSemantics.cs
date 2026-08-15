namespace OfficeIMO.Pdf;

/// <summary>Owns optional image-mask dictionary semantics shared by image normalization paths.</summary>
internal static class PdfImageMaskSemantics {
    internal static bool HasSoftMask(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!dictionary.Items.TryGetValue("SMask", out PdfObject? softMaskObject)) return false;
        if (!PdfObjectLookup.TryResolveReferenceChain(objects, softMaskObject, out PdfObject? resolved)) return true;
        return resolved is not PdfNull &&
            (resolved is not PdfName softMaskName || !string.Equals(softMaskName.Name, "None", StringComparison.Ordinal));
    }

    internal static bool HasUnsupportedSoftMaskMatte(
        PdfDictionary dictionary,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!dictionary.Items.TryGetValue("SMask", out PdfObject? softMaskObject) ||
            !PdfObjectLookup.TryResolveReferenceChain(objects, softMaskObject, out PdfObject? resolvedSoftMask) ||
            resolvedSoftMask is not PdfStream softMask ||
            !softMask.Dictionary.Items.TryGetValue("Matte", out PdfObject? matteObject)) {
            return false;
        }

        return !PdfObjectLookup.TryResolveReferenceChain(objects, matteObject, out PdfObject? resolvedMatte) ||
            resolvedMatte is not PdfNull;
    }
}

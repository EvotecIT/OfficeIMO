namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private PdfDictionary? ResolveDict(PdfObject? obj) {
        if (obj is null) return null;
        if (obj is PdfDictionary d) return d;
        if (obj is PdfReference r && PdfObjectLookup.TryGet(_objects, r, out var ind) && ind.Value is PdfDictionary dd) return dd;
        return null;
    }

    private PdfObject? ResolveObject(PdfObject? obj) {
        return PdfObjectLookup.Resolve(_objects, obj);
    }

    private PdfArray? ResolveArray(PdfObject? obj) {
        if (obj is null) return null;
        if (obj is PdfArray a) return a;
        if (obj is PdfReference r && PdfObjectLookup.TryGet(_objects, r, out var ind) && ind.Value is PdfArray aa) return aa;
        return null;
    }

    private int FindExactObjectNumberFor(PdfDictionary dict) {
        foreach (var kv in _objects) if (ReferenceEquals(kv.Value.Value, dict)) return kv.Key;
        return 0;
    }

    private int FindObjectNumberFor(PdfDictionary dict) {
        foreach (var kv in _objects) if (ReferenceEquals(kv.Value.Value, dict)) return kv.Key;
        // As a fallback when dictionary was re-parsed separately, match by identity via a simple scan of Page objects
        foreach (var kv in _objects) if (kv.Value.Value is PdfDictionary d && d.Get<PdfName>("Type")?.Name == "Page") return kv.Key;
        return 0;
    }
}

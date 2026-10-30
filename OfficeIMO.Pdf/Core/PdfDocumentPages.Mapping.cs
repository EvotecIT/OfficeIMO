namespace OfficeIMO.Pdf;

/// <remarks>Geometry uses the source page's default user-space units: one unit is /UserUnit physical points (1/72 inch each), with /UserUnit defaulting to 1. Coordinates are not automatically normalized to physical points. Explicitly named visual bounds use their documented visual coordinate system.</remarks>
public sealed partial class PdfDocumentPages {
    /// <summary>Reorders every page and reports the object numbers needed to retain annotation or object selections.</summary>
    public PdfPageRewriteResult ReorderWithMapping(params int[] pageNumbers) {
        byte[] input = _document.GetBytesForOperation();
        var result = PdfPageEditor.ReorderPagesWithMapping(input, _document.ReadOptions, pageNumbers);
        PdfDocument rewritten = _document.WithBytes(input, result.Bytes);
        return new PdfPageRewriteResult(result.Bytes, result.Map, rewritten.ReadOptions);
    }
}

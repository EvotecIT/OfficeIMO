namespace OfficeIMO.Pdf;

/// <remarks>Native page coordinates use default user-space units: one unit is /UserUnit physical points (1/72 inch each), defaulting to 1. Local text/font and matrix values retain their native coordinate spaces. Explicit visual projections use their documented coordinates.</remarks>
public sealed partial class PdfDocumentPages {
    /// <summary>Reorders every page and reports the object numbers needed to retain annotation or object selections.</summary>
    public PdfPageRewriteResult ReorderWithMapping(params int[] pageNumbers) {
        byte[] input = _document.GetBytesForOperation();
        var result = PdfPageEditor.ReorderPagesWithMapping(input, _document.ReadOptions, pageNumbers);
        PdfDocument rewritten = _document.WithBytes(input, result.Bytes);
        return new PdfPageRewriteResult(result.Bytes, result.Map, rewritten.ReadOptions);
    }
}

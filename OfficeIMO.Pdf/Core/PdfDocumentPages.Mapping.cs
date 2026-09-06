namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages {
    /// <summary>Reorders every page and reports the object numbers needed to retain annotation or object selections.</summary>
    public PdfPageRewriteResult ReorderWithMapping(params int[] pageNumbers) {
        byte[] input = _document.GetBytesForOperation();
        var result = PdfPageEditor.ReorderPagesWithMapping(input, _document.ReadOptions, pageNumbers);
        PdfDocument rewritten = _document.WithBytes(input, result.Bytes);
        return new PdfPageRewriteResult(result.Bytes, result.Map, rewritten.ReadOptions);
    }
}

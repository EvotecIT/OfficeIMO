namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Flattens readable AcroForm fields and supported visual annotations into one new PDF artifact.</summary>
    public PdfInteractiveContentFlattenResult FlattenInteractiveContent(PdfInteractiveContentFlattenOptions? options = null) =>
        PdfInteractiveContentFlattener.Flatten(GetBytesForOperation(), options, ReadOptions);
}

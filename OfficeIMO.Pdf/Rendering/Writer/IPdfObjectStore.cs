namespace OfficeIMO.Pdf;

internal interface IPdfObjectStore :
    System.Collections.Generic.IList<byte[]>,
    System.Collections.Generic.IReadOnlyList<byte[]>,
    System.IDisposable {
    PdfVisualResourceCache VisualResources { get; }
}

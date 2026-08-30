namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal PdfDictionary? GetFontInspectionResources() =>
        ResolveDictionary(GetInheritedValue("Resources"));
}

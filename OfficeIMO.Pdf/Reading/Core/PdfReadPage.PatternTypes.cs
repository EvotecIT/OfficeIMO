namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private bool HasSupportedOptionalPatternType(PdfDictionary dictionary) {
        // ISO 32000-1 tables 75/76 make Type optional. PDF null dictionary
        // entries are absent; a present, resolved name must still be Pattern.
        return !dictionary.Items.TryGetValue("Type", out PdfObject? value) ||
               ResolveEffectObject(value) is PdfNull or PdfName { Name: "Pattern" };
    }
}

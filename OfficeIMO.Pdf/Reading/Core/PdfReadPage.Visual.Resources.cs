namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private bool TryResolveStrictResources(
        PdfDictionary owner,
        PdfDictionary? inheritedResources,
        out PdfDictionary? resources) {
        resources = inheritedResources;
        if (!owner.Items.TryGetValue("Resources", out PdfObject? resourceObject)) return true;

        PdfObject? resolved = ResolveEffectObject(resourceObject);
        if (resolved is PdfNull) return true;
        if (resolved is PdfDictionary dictionary) {
            resources = dictionary;
            return true;
        }

        return false;
    }
}

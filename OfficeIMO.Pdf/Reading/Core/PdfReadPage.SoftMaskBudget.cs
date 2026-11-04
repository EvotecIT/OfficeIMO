namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal sealed partial class PageContentBudget {
        // Isolated paint groups have independent graphics state, but share recursion ancestry.
        internal HashSet<PdfStream> ActiveSoftMasks { get; } = new();
    }
}

using System.Collections.Generic;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static bool TryResolveReorderedLogicalText(IEnumerable<HtmlRenderVisual> visuals, out string logicalText) =>
        HtmlRenderLogicalText.TryResolveReorderedText(visuals, out logicalText);
}

using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static bool HasCanvasContent(IReadOnlyList<PdfCore.PdfCanvasItem> items) {
        foreach (PdfCore.PdfCanvasItem item in items) {
            bool hasContent = item switch {
                PdfCore.PdfCanvasOutlineItem => false,
                PdfCore.PdfCanvasNamedDestinationItem => false,
                PdfCore.PdfCanvasNamedDestinationLinkItem => false,
                PdfCore.PdfCanvasArtifactItem artifact => HasCanvasContent(artifact.Items),
                PdfCore.PdfCanvasFigureItem figure => HasCanvasContent(figure.Items),
                PdfCore.PdfCanvasStructureItem structure => HasCanvasContent(structure.Items),
                PdfCore.PdfCanvasActualTextItem actualText => HasCanvasContent(actualText.Items),
                PdfCore.PdfCanvasClipItem clip => HasCanvasContent(clip.Items),
                PdfCore.PdfCanvasEffectItem effect => HasCanvasContent(effect.Items),
                _ => true
            };
            if (hasContent) return true;
        }
        return false;
    }
}

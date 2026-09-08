namespace OfficeIMO.Pdf;

public static partial class PdfAdvancedUnderstandingStages {
    private static PdfLogicalVisualBounds GetSelectionBounds(PdfUnderstandingPageContext context,
        double x, double y, double advance, double height, double angle) {
        double radians = angle * Math.PI / 180D, cos = Math.Cos(radians), sin = Math.Sin(radians);
        double endX = x + cos * advance, endY = y + sin * advance;
        double upX = -sin * height, upY = cos * height;
        PdfVisualBounds visual = context.ToVisualBounds(
            Math.Min(Math.Min(x, endX), Math.Min(x + upX, endX + upX)),
            Math.Min(Math.Min(y, endY), Math.Min(y + upY, endY + upY)),
            Math.Max(Math.Max(x, endX), Math.Max(x + upX, endX + upX)),
            Math.Max(Math.Max(y, endY), Math.Max(y + upY, endY + upY)));
        return new PdfLogicalVisualBounds(visual.Left, visual.Top, visual.Right, visual.Bottom);
    }
}

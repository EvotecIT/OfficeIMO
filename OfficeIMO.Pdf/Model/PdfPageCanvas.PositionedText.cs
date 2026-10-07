namespace OfficeIMO.Pdf;

public sealed partial class PdfPageCanvas {
    // Adapter-owned text has already been broken into lines. Preserve its exact whitespace
    // and let the adapter's enclosing clips govern paint, rather than clipping glyph ink
    // to an advance/line-height layout box.
    internal PdfPageCanvas PositionedText(IEnumerable<PdfTextRun> runs, PdfCanvasTextStructureRole structureRole,
        double x, double y, double width, double height, PdfColor? defaultColor = null,
        PdfAlign align = PdfAlign.Left, double? fontSize = null, double? lineHeight = null,
        double? advanceWidth = null, double clipTopOverflow = 0D, double fontMetricScale = 1D) {
        if (fontMetricScale <= 0D || double.IsNaN(fontMetricScale) || double.IsInfinity(fontMetricScale))
            throw new ArgumentOutOfRangeException(nameof(fontMetricScale));
        AddText(runs, structureRole, x, y, width, height, defaultColor, align, fontSize, lineHeight);
        ((PdfCanvasTextItem)_items[_items.Count - 1]).PreservePositionedText = true;
        ((PdfCanvasTextItem)_items[_items.Count - 1]).PositionedAdvanceWidth = advanceWidth;
        ((PdfCanvasTextItem)_items[_items.Count - 1]).PositionedClipTopOverflow = clipTopOverflow;
        ((PdfCanvasTextItem)_items[_items.Count - 1]).PositionedFontMetricScale = fontMetricScale;
        return this;
    }
}

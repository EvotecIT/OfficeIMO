namespace OfficeIMO.Pdf;

public sealed partial class PdfPageCanvas {
    // Adapter-owned text has already been broken into lines. Preserve its exact whitespace
    // and let the adapter's enclosing clips govern paint, rather than clipping glyph ink
    // to an advance/line-height layout box.
    internal PdfPageCanvas PositionedText(IEnumerable<PdfTextRun> runs, PdfCanvasTextStructureRole structureRole,
        double x, double y, double width, double height, PdfColor? defaultColor = null,
        PdfAlign align = PdfAlign.Left, double? fontSize = null, double? lineHeight = null, double? advanceWidth = null) {
        AddText(runs, structureRole, x, y, width, height, defaultColor, align, fontSize, lineHeight);
        ((PdfCanvasTextItem)_items[_items.Count - 1]).PreservePositionedText = true;
        ((PdfCanvasTextItem)_items[_items.Count - 1]).PositionedAdvanceWidth = advanceWidth;
        return this;
    }
}

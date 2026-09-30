using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    internal bool HasTracking => _tracking != null;

    internal PdfTextShowCommand EncodeTextShowCommand(string text) =>
        ToTextShowCommand(text, ShapeText(text));

    internal double MeasureShapedTextWidth(string text, PdfGlyphRun run, double fontSize, double fontMetricScale = 1D) {
        double nominal = run.TotalAdvanceWidth1000 * fontSize / 1000D;
        if (_tracking == null || run.Direction == OfficeTextDirection.TopToBottom) return nominal;
        int count = 0;
        foreach (bool boundary in GetTrackingBoundaries(text, run)) if (boundary) count++;
        double adjustment = _tracking.GetAdjustment(fontSize / fontMetricScale) * fontSize / UnitsPerEm;
        return Math.Abs(nominal + (run.TotalAdvanceWidth1000 < 0 ? -adjustment : adjustment) * count);
    }

    internal PdfTextShowCommand ToTextShowCommand(string text, PdfGlyphRun run, double fontMetricScale = 1D) {
        PdfTextShowCommand command = run.ToTextShowCommand();
        if (_tracking == null) return command;
        return new PdfTextShowCommand(command.GlyphHex, run.Glyphs, command.ActualText,
            logicalGlyphs: command.LogicalGlyphs, advanceWidth1000: command.AdvanceWidth1000, wordSpaceCount: command.WordSpaceCount, visualGlyphs: command.VisualGlyphs,
            tracking: _tracking, unitsPerEm: UnitsPerEm, trackingBoundaries: GetTrackingBoundaries(text, run), negativeTracking: run.TotalAdvanceWidth1000 < 0, fontMetricScale: fontMetricScale);
    }

    private static bool[] GetTrackingBoundaries(string text, PdfGlyphRun run) {
        var indexes = new int[run.Glyphs.Count];
        for (int index = 0; index < indexes.Length; index++) indexes[index] = run.Glyphs[index].TextIndex;
        // PDF TJ advances happen after painting, for either sign. Core's negative contour
        // cursor advances before painting, so its first-glyph convention does not apply here.
        return OfficeOpenTypeTracking.GetBoundaries(text, indexes);
    }
}

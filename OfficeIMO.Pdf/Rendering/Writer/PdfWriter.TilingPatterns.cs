using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed class PageTilingPattern {
        internal PageTilingPattern(OfficeDrawingTilingPattern pattern, double originX, double originTopY) {
            Pattern = pattern;
            OriginX = originX;
            OriginTopY = originTopY;
        }
        internal OfficeDrawingTilingPattern Pattern { get; }
        internal double OriginX { get; }
        internal double OriginTopY { get; }
    }

    // The existing effect traversal resolves child Forms and images before this
    // helper. A small Form owns the Pattern resource so its matrix follows the
    // containing Form/CTM, including nested transformed SVG viewports.
    private static int AddTilingPatternForm(IList<byte[]> objects, PageEffectGroup effect, string tileContent,
        List<(string Name, int Id)> fonts, List<(string Name, int Id)> xobjects,
        List<(string Name, int Id)> graphicsStates, List<(string Name, int Id)> shadings) {
        PageTilingPattern capture = effect.TilingPattern!;
        OfficeDrawingTilingPattern pattern = capture.Pattern;
        double width = pattern.InnerTile.Width, height = pattern.InnerTile.Height;
        OfficeTransform matrix = OfficeTransform.Scale(1D, -1D)
            .Then(OfficeTransform.Translate(pattern.OriginX, pattern.OriginY + height))
            .Then(pattern.Transform)
            .Then(new OfficeTransform(1D, 0D, 0D, -1D, capture.OriginX, capture.OriginTopY));
        pattern.Transform.TryInvert(out OfficeTransform inverse);
        var bounds = inverse.TransformRectangleBounds(pattern.Area.X, pattern.Area.Y, pattern.Area.Width, pattern.Area.Height);
        // A non-repeating axis uses a cell stride larger than either distance to
        // the visible area; adjacent cells therefore cannot enter its clip.
        double stepX = pattern.RepeatX ? pattern.HorizontalStep : NonRepeatingPatternStep(bounds.Left, bounds.Right, pattern.OriginX, width);
        double stepY = pattern.RepeatY ? pattern.VerticalStep : NonRepeatingPatternStep(bounds.Top, bounds.Bottom, pattern.OriginY, height);
        var resources = new StringBuilder();
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "Font", FilterPdfResources(tileContent, fonts));
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "XObject", FilterPdfResources(tileContent, xobjects));
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "ExtGState", FilterPdfResources(tileContent, graphicsStates));
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "Shading", FilterPdfResources(tileContent, shadings));
        string entries = "/Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 " +
            PdfNumberFormatter.Precise(width) + " " + PdfNumberFormatter.Precise(height) + "] /XStep " +
            PdfNumberFormatter.Precise(stepX) + " /YStep " + PdfNumberFormatter.Precise(stepY) +
            " /Matrix [" + FormatPatternMatrix(matrix) + "] /Resources <<" + resources + " >>";
        int patternId = AddFlateStreamObject(objects, PdfEncoding.Latin1GetBytes(tileContent), entries);
        var buffer = new StringBuilder("/Pattern cs /Pt scn\n");
        new ContentStreamBuilder(buffer).Rectangle(effect.BoundsLeft, effect.BoundsBottom,
            effect.BoundsRight - effect.BoundsLeft, effect.BoundsTop - effect.BoundsBottom, preciseCoordinates: true).FillPath();
        string formEntries = "/Type /XObject /Subtype /Form /FormType 1 /BBox [" +
            PdfNumberFormatter.Precise(effect.BoundsLeft) + " " + PdfNumberFormatter.Precise(effect.BoundsBottom) + " " +
            PdfNumberFormatter.Precise(effect.BoundsRight) + " " + PdfNumberFormatter.Precise(effect.BoundsTop) + "]" +
            " /Group << /S /Transparency /I true /K false >> /Resources << /Pattern << /Pt " + patternId + " 0 R >> >>";
        return AddFlateStreamObject(objects, Encoding.ASCII.GetBytes(buffer.ToString()), formEntries);
    }

    private static double NonRepeatingPatternStep(double minimum, double maximum, double origin, double size) =>
        2D * (Math.Max(Math.Abs(minimum - origin), Math.Abs(maximum - origin)) + size + 1D);

    private static string FormatPatternMatrix(OfficeTransform matrix) => string.Join(" ",
        new[] { matrix.M11, matrix.M12, matrix.M21, matrix.M22, matrix.OffsetX, matrix.OffsetY }.Select(PdfNumberFormatter.Precise));
}

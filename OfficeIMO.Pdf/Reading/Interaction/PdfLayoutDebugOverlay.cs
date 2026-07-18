using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Creates dependency-free word, line, region, and reading-order overlays through OfficeIMO.Drawing.</summary>
internal static class PdfLayoutDebugOverlay {
    /// <summary>Builds a transparent vector overlay for one page in rendered top-left coordinates.</summary>
    public static OfficeDrawing CreateDrawing(
        byte[] pdf,
        int pageNumber,
        PdfLayoutDebugOverlayOptions? options = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfLayoutDebugOverlayOptions effective = options ?? new PdfLayoutDebugOverlayOptions();
        if (effective.MaxElements <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "Maximum overlay elements must be positive.");
        }

        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        if (pageNumber <= 0 || pageNumber > document.Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), "Page number is outside the PDF page count.");
        }

        PdfReadPage page = document.Pages[pageNumber - 1];
        PdfPageInteractionMap interactions = PdfPageInteractionMap.Create(pdf, pageNumber, readOptions: readOptions);
        StructuredPage structure = page.ExtractStructured(layoutOptions);
        var drawing = new OfficeDrawing(interactions.Width, interactions.Height);
        int elementCount = 0;
        if (effective.ShowRegions) {
            AddRegions(drawing, page, structure, effective, ref elementCount);
        }

        if (effective.ShowLines || effective.ShowReadingOrder) {
            AddLines(drawing, page, structure, effective, ref elementCount);
        }

        if (effective.ShowWords) {
            AddWords(drawing, interactions.TextRegions, effective, ref elementCount);
        }

        return drawing;
    }

    /// <summary>Creates an SVG debug overlay for one page.</summary>
    public static string ToSvg(
        byte[] pdf,
        int pageNumber,
        PdfLayoutDebugOverlayOptions? options = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        return OfficeDrawingSvgExporter.ToSvg(CreateDrawing(pdf, pageNumber, options, layoutOptions, readOptions));
    }

    /// <summary>Creates a transparent PNG debug overlay for one page.</summary>
    public static byte[] ToPng(
        byte[] pdf,
        int pageNumber,
        double scale = 1D,
        PdfLayoutDebugOverlayOptions? options = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        return OfficeDrawingRasterRenderer.ToPng(
            CreateDrawing(pdf, pageNumber, options, layoutOptions, readOptions),
            scale);
    }

    private static void AddWords(
        OfficeDrawing drawing,
        IReadOnlyList<PdfPageInteractionRegion> textRegions,
        PdfLayoutDebugOverlayOptions options,
        ref int elementCount) {
        var word = new List<PdfPageInteractionRegion>();
        for (int i = 0; i <= textRegions.Count; i++) {
            PdfPageInteractionRegion? region = i < textRegions.Count ? textRegions[i] : null;
            if (region is not null && !string.IsNullOrWhiteSpace(region.Text)) {
                word.Add(region);
                continue;
            }

            if (word.Count > 0) {
                AddBox(drawing, Bounds(word.Select(static item => item.Quad)), options.WordColor, 0.65D, options, ref elementCount);
                word.Clear();
            }
        }
    }

    private static void AddLines(
        OfficeDrawing drawing,
        PdfReadPage page,
        StructuredPage structure,
        PdfLayoutDebugOverlayOptions options,
        ref int elementCount) {
        for (int i = 0; i < structure.LinesDetailed.Count; i++) {
            StructuredLine line = structure.LinesDetailed[i];
            DebugBounds bounds = TransformBounds(
                page,
                line.XStart,
                line.Y - Math.Max(1D, line.FontSize * 0.2D),
                line.XEnd,
                line.Y + Math.Max(1D, line.FontSize),
                drawing.Width,
                drawing.Height);
            if (options.ShowLines) {
                AddBox(drawing, bounds, options.LineColor, 1D, options, ref elementCount);
            }

            if (options.ShowReadingOrder && bounds.Width > 0D && bounds.Height > 0D) {
                EnsureCapacity(options, ++elementCount);
                double labelWidth = Math.Min(24D, Math.Max(10D, bounds.Width));
                double labelHeight = Math.Min(12D, Math.Max(8D, bounds.Height));
                drawing.AddText(
                    (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture),
                    bounds.Left,
                    bounds.Top,
                    labelWidth,
                    labelHeight,
                    new OfficeFontInfo("Arial", 7D),
                    options.LineColor,
                    wrapText: false);
            }
        }
    }

    private static void AddRegions(
        OfficeDrawing drawing,
        PdfReadPage page,
        StructuredPage structure,
        PdfLayoutDebugOverlayOptions options,
        ref int elementCount) {
        for (int i = 0; i < structure.Paragraphs.Count; i++) {
            StructuredParagraph paragraph = structure.Paragraphs[i];
            double largestFont = paragraph.Lines.Count == 0
                ? 10D
                : paragraph.Lines.Max(static line => Math.Max(1D, line.FontSize));
            DebugBounds bounds = TransformBounds(
                page,
                paragraph.XStart,
                paragraph.YBottom - largestFont * 0.2D,
                paragraph.XEnd,
                paragraph.YTop + largestFont,
                drawing.Width,
                drawing.Height);
            AddBox(drawing, bounds, options.RegionColor, 1.4D, options, ref elementCount);
        }
    }

    private static void AddBox(
        OfficeDrawing drawing,
        DebugBounds bounds,
        OfficeColor color,
        double strokeWidth,
        PdfLayoutDebugOverlayOptions options,
        ref int elementCount) {
        if (bounds.Width <= 0.01D || bounds.Height <= 0.01D) {
            return;
        }

        EnsureCapacity(options, ++elementCount);
        OfficeShape shape = OfficeShape.Rectangle(bounds.Width, bounds.Height);
        shape.StrokeColor = color;
        shape.StrokeWidth = strokeWidth;
        shape.FillColor = null;
        drawing.AddShape(shape, bounds.Left, bounds.Top);
    }

    private static DebugBounds Bounds(IEnumerable<PdfSelectionQuad> quads) {
        double left = double.MaxValue;
        double top = double.MaxValue;
        double right = double.MinValue;
        double bottom = double.MinValue;
        foreach (PdfSelectionQuad quad in quads) {
            left = Math.Min(left, quad.Left);
            top = Math.Min(top, quad.Top);
            right = Math.Max(right, quad.Right);
            bottom = Math.Max(bottom, quad.Bottom);
        }

        return left == double.MaxValue
            ? new DebugBounds(0D, 0D, 0D, 0D)
            : new DebugBounds(left, top, right, bottom);
    }

    private static DebugBounds TransformBounds(
        PdfReadPage page,
        double left,
        double bottom,
        double right,
        double top,
        double pageWidth,
        double pageHeight) {
        (double X, double Y) p1 = page.TransformPointToVisual(left, bottom);
        (double X, double Y) p2 = page.TransformPointToVisual(right, bottom);
        (double X, double Y) p3 = page.TransformPointToVisual(right, top);
        (double X, double Y) p4 = page.TransformPointToVisual(left, top);
        double visualLeft = Math.Max(0D, Math.Min(Math.Min(p1.X, p2.X), Math.Min(p3.X, p4.X)));
        double visualRight = Math.Min(pageWidth, Math.Max(Math.Max(p1.X, p2.X), Math.Max(p3.X, p4.X)));
        double visualTop = Math.Max(0D, pageHeight - Math.Max(Math.Max(p1.Y, p2.Y), Math.Max(p3.Y, p4.Y)));
        double visualBottom = Math.Min(pageHeight, pageHeight - Math.Min(Math.Min(p1.Y, p2.Y), Math.Min(p3.Y, p4.Y)));
        return new DebugBounds(visualLeft, visualTop, visualRight, visualBottom);
    }

    private static void EnsureCapacity(PdfLayoutDebugOverlayOptions options, int actual) {
        if (actual > options.MaxElements) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.InteractionRegions, options.MaxElements, actual);
        }
    }

    private readonly struct DebugBounds {
        internal DebugBounds(double left, double top, double right, double bottom) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        internal double Left { get; }
        internal double Top { get; }
        internal double Right { get; }
        internal double Bottom { get; }
        internal double Width => Math.Max(0D, Right - Left);
        internal double Height => Math.Max(0D, Bottom - Top);
    }
}

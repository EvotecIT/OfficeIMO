using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// A bounded vector drawing on a page, typically a chart: many painted paths with short labels inside or around
/// it. Editable converters cannot rebuild it as native content, so they place it as a picture of the region and
/// leave out the labels it already shows.
/// </summary>
internal sealed class PdfLogicalVectorFigure {
    private readonly Func<double, byte[]?> _render;

    internal PdfLogicalVectorFigure(double left, double top, double right, double bottom,
        IReadOnlyList<PdfLogicalTextBlock> labels, int primitiveCount, Func<double, byte[]?> render) {
        Left = left; Top = top; Right = right; Bottom = bottom; Labels = labels; PrimitiveCount = primitiveCount; _render = render;
    }

    /// <summary>Visual (top-left origin) bounds in points.</summary>
    internal double Left { get; }
    internal double Top { get; }
    internal double Right { get; }
    internal double Bottom { get; }
    internal double Width => Right - Left;
    internal double Height => Bottom - Top;

    /// <summary>Text blocks drawn as part of the figure (axis, legend and data labels).</summary>
    internal IReadOnlyList<PdfLogicalTextBlock> Labels { get; }

    internal int PrimitiveCount { get; }

    /// <summary>Renders the figure region to PNG at the given scale, or returns null when it cannot be rendered.</summary>
    internal byte[]? RenderPng(double scale) => _render(scale);
}

internal static partial class PdfLogicalVectorFigureAnalysis {
    private const int MinimumPrimitives = 6;
    private const int MaximumPrimitives = 20_000;

    private readonly struct Box {
        internal Box(double left, double top, double right, double bottom, bool filled, bool stroked, bool rectangle, double paintOrder) {
            Left = left; Top = top; Right = right; Bottom = bottom; Filled = filled;
            Stroked = stroked; Rectangle = rectangle;
            PaintOrder = paintOrder;
        }
        internal double Left { get; }
        internal double Top { get; }
        internal double Right { get; }
        internal double Bottom { get; }
        internal bool Filled { get; }
        internal bool Stroked { get; }
        internal bool Rectangle { get; }
        internal double PaintOrder { get; }
        internal bool Touches(Box other, double gap) =>
            Left - gap <= other.Right && other.Left - gap <= Right && Top - gap <= other.Bottom && other.Top - gap <= Bottom;
    }

    internal static IReadOnlyList<PdfLogicalVectorFigure> Detect(
        PdfReadPage page,
        IReadOnlyList<PdfLogicalTextBlock> textBlocks,
        IReadOnlyList<PdfLogicalTable> tables,
        IReadOnlyList<PdfImagePlacement> imagePlacements,
        CancellationToken cancellationToken) {
        (double pageWidth, double pageHeight) = page.GetVisualPageSize();
        double pageArea = Math.Max(1D, pageWidth * pageHeight);
        IReadOnlyList<PdfPageVisualPrimitive> primitives;
        try {
            primitives = page.GetIdentityVisualPrimitives(cancellationToken);
        } catch (Exception exception) when (exception is not OperationCanceledException && exception is not PdfReadLimitException) {
            return Array.Empty<PdfLogicalVectorFigure>();
        }
        if (primitives.Count < MinimumPrimitives || primitives.Count > MaximumPrimitives) return Array.Empty<PdfLogicalVectorFigure>();

        PdfVisualBounds[] tableBounds = tables
            .Where(static table => table.Columns.Count > 0)
            .Select(table => GetTableVisualBounds(page, table))
            .ToArray();
        PdfVisualBounds[] imageBounds = imagePlacements
            .Select(placement => page.TransformBoundsToVisual(placement.X, placement.Y, placement.X + placement.Width, placement.Y + placement.Height))
            .ToArray();

        var boxes = new List<Box>();
        var tableEvidence = new List<Box>();
        foreach (PdfPageVisualPrimitive primitive in primitives) {
            cancellationToken.ThrowIfCancellationRequested();
            if (primitive.Kind is not (PdfPageVisualPrimitiveKind.Rectangle or PdfPageVisualPrimitiveKind.Path or PdfPageVisualPrimitiveKind.Line)) continue;
            bool filled = primitive.FillColor.HasValue && primitive.FillColor.Value.A > 0 || primitive.FillGradient != null || primitive.FillRadialGradient != null;
            bool stroked = primitive.StrokeColor.HasValue && primitive.StrokeColor.Value.A > 0;
            if (!filled && !stroked) continue;
            double left, top, right, bottom;
            if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
                left = Math.Min(primitive.X1, primitive.X2); right = Math.Max(primitive.X1, primitive.X2);
                top = Math.Min(primitive.Y1, primitive.Y2); bottom = Math.Max(primitive.Y1, primitive.Y2);
            } else {
                left = Math.Min(primitive.X, primitive.X + primitive.Width); right = Math.Max(primitive.X, primitive.X + primitive.Width);
                top = Math.Min(primitive.Y, primitive.Y + primitive.Height); bottom = Math.Max(primitive.Y, primitive.Y + primitive.Height);
            }
            double area = (right - left) * (bottom - top);
            // Page backgrounds and page-wide rules frame content rather than draw a figure.
            if (area >= pageArea * 0.5D || right - left >= pageWidth * 0.9D) continue;
            var box = new Box(left, top, right, bottom, filled, stroked,
                primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle, primitive.PaintOrder);
            tableEvidence.Add(box);
            double centerX = (left + right) / 2D, centerY = (top + bottom) / 2D;
            if (tableBounds.Any(bounds => centerX >= bounds.Left - 4D && centerX <= bounds.Right + 4D && centerY >= bounds.Top - 4D && centerY <= bounds.Bottom + 4D)) continue;
            boxes.Add(box);
        }
        if (boxes.Count < MinimumPrimitives || tableEvidence.Count > 3000) return Array.Empty<PdfLogicalVectorFigure>();

        // Group primitives that touch or nearly touch.
        int[] parent = Enumerable.Range(0, boxes.Count).ToArray();
        for (int first = 0; first < boxes.Count; first++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int second = first + 1; second < boxes.Count; second++) {
                if (boxes[first].Touches(boxes[second], 6D)) Union(parent, first, second);
            }
        }

        var figures = new List<PdfLogicalVectorFigure>();
        var claimedLabels = new HashSet<PdfLogicalTextBlock>();
        foreach (IGrouping<int, int> group in Enumerable.Range(0, boxes.Count).GroupBy(index => Find(parent, index))) {
            Box[] members = group.Select(index => boxes[index]).ToArray();
            if (members.Length < MinimumPrimitives) continue;
            // Cell paint and connected rulings are table evidence even when the
            // text detector recovered only part of the grid. Keep its text editable.
            if (IsRuledCellGrid(members, tableEvidence, cancellationToken)) continue;
            double left = members.Min(static box => box.Left), top = members.Min(static box => box.Top);
            double right = members.Max(static box => box.Right), bottom = members.Max(static box => box.Bottom);
            if (right - left < 60D || bottom - top < 40D) continue;
            // A chart has several marks: bars, slices, points or grid lines.
            int marks = members.Count(box => box.Filled && (box.Right - box.Left) * (box.Bottom - box.Top) < (right - left) * (bottom - top) * 0.6D);
            int lines = members.Count(static box => !box.Filled);
            if (marks < 3 && lines < 4) continue;
            if (imageBounds.Any(image => Overlap(image, left, top, right, bottom) > 0.5D)) continue;

            var inside = new List<PdfLogicalTextBlock>();
            bool hasProse = false;
            foreach (PdfLogicalTextBlock block in textBlocks) {
                if (!TryGetVisualBounds(page, block, out PdfVisualBounds bounds)) continue;
                double centerX = (bounds.Left + bounds.Right) / 2D, centerY = (bounds.Top + bounds.Bottom) / 2D;
                double margin = Math.Max(4D, block.FontSize * 1.8D);
                bool within = centerX >= left - margin && centerX <= right + margin && centerY >= top - margin && centerY <= bottom + margin;
                if (!within) continue;
                if (CountWords(block.Text) > 6) { hasProse = true; break; }
                inside.Add(block);
            }
            // Text panels and cards hold sentences; charts hold short labels.
            if (hasProse || inside.Count == 0 && marks < 3) continue;
            foreach (PdfLogicalTextBlock label in inside) {
                if (!claimedLabels.Add(label)) continue;
                if (!TryGetVisualBounds(page, label, out PdfVisualBounds bounds)) continue;
                left = Math.Min(left, bounds.Left); top = Math.Min(top, bounds.Top);
                right = Math.Max(right, bounds.Right); bottom = Math.Max(bottom, bounds.Bottom);
            }
            double padding = 4D;
            double figureLeft = Math.Max(0D, left - padding), figureTop = Math.Max(0D, top - padding);
            double figureRight = Math.Min(pageWidth, right + padding), figureBottom = Math.Min(pageHeight, bottom + padding);
            figures.Add(new PdfLogicalVectorFigure(figureLeft, figureTop, figureRight, figureBottom, inside.AsReadOnly(), members.Length,
                scale => Render(page, figureLeft, figureTop, figureRight - figureLeft, figureBottom - figureTop, scale)));
        }
        return figures.Count == 0 ? Array.Empty<PdfLogicalVectorFigure>() : figures.AsReadOnly();
    }

    private static byte[]? Render(PdfReadPage page, double left, double top, double width, double height, double scale) {
        try {
            OfficeDrawing drawing = page.ToDrawing(CancellationToken.None, functionShadingScale: scale);
            var region = new OfficeDrawing(width, height);
            region.AddDrawingForClippedRendering(drawing, -left, -top, null);
            return OfficeDrawingRasterRenderer.ToPng(region, new OfficeDrawingRasterRenderOptions {
                Scale = scale,
                Background = OfficeColor.White,
                PdfRasterConventions = true
            });
        } catch (Exception exception) when (exception is not OperationCanceledException) {
            return null;
        }
    }

    private static double Overlap(PdfVisualBounds image, double left, double top, double right, double bottom) {
        double width = Math.Min(image.Right, right) - Math.Max(image.Left, left);
        double height = Math.Min(image.Bottom, bottom) - Math.Max(image.Top, top);
        double area = Math.Max(1D, (right - left) * (bottom - top));
        return width <= 0D || height <= 0D ? 0D : width * height / area;
    }

    private static int CountWords(string text) =>
        text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;

    private static bool TryGetVisualBounds(PdfReadPage page, PdfLogicalTextBlock block, out PdfVisualBounds bounds) {
        if (block.VisualBounds is PdfLogicalVisualBounds visual) {
            bounds = new PdfVisualBounds(visual.Left, visual.Top, visual.Right, visual.Bottom);
            return visual.Right > visual.Left;
        }
        if (block.XEnd <= block.XStart) { bounds = default; return false; }
        bounds = page.TransformBoundsToVisual(block.XStart, block.BaselineY - block.FontSize * 0.25D, block.XEnd, block.BaselineY + block.FontSize);
        return true;
    }

    private static PdfVisualBounds GetTableVisualBounds(PdfReadPage page, PdfLogicalTable table) {
        if (table.VisualBounds is PdfLogicalVisualBounds visual) return new PdfVisualBounds(visual.Left, visual.Top, visual.Right, visual.Bottom);
        double left = table.Columns.Min(static column => Math.Min(column.From, column.To));
        double right = table.Columns.Max(static column => Math.Max(column.From, column.To));
        return page.TransformBoundsToVisual(left, Math.Min(table.YBottom, table.YTop) - 6D, right, Math.Max(table.YBottom, table.YTop) + 12D);
    }

    private static int Find(int[] parent, int value) {
        while (parent[value] != value) {
            parent[value] = parent[parent[value]];
            value = parent[value];
        }
        return value;
    }

    private static void Union(int[] parent, int left, int right) {
        int leftRoot = Find(parent, left), rightRoot = Find(parent, right);
        if (leftRoot != rightRoot) parent[rightRoot] = leftRoot;
    }
}

using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    // This small ordinary-grid profile supplements text-derived boundaries. Large
    // or complex pages keep the existing inference without retaining parser data.
    private const int MaximumTableStrokeSamples = 256;
    private const double TableStrokeTolerance = 0.75D;

    private int CountWithPhysicalTableFrames(List<TableStrokeSample> samples,
        IReadOnlyList<PdfStructuredTable> tables, int inferredUnrepresented) {
        List<PhysicalTableStrokeFrame> frames = FindPhysicalTableStrokeFrames(samples, tables);
        foreach (TableStrokeSample sample in samples) {
            bool inside = false, represented = false;
            foreach (PhysicalTableStrokeFrame frame in frames) {
                if (!frame.Contains(sample)) continue;
                inside = true;
                represented |= frame.IsBorder(sample);
            }
            if (!inside || represented == sample.InferredBorder) continue;
            inferredUnrepresented += represented ? -1 : 1;
        }
        return inferredUnrepresented;
    }

    /// <summary>
    /// Qualifies adjoining, equally painted row rectangles containing every detected
    /// row. Interior dividers must recur across every row; a lone stroke cannot
    /// establish a column boundary merely by being near narrow text extents.
    /// </summary>
    private List<PhysicalTableStrokeFrame> FindPhysicalTableStrokeFrames(
        List<TableStrokeSample> samples, IReadOnlyList<PdfStructuredTable> tables) {
        var result = new List<PhysicalTableStrokeFrame>();
        TableStrokeSample[] rectangles = samples.Where(sample => sample.Ordinary &&
            sample.Kind == PdfPageVisualPrimitiveKind.Rectangle).ToArray();
        foreach (PdfStructuredTable table in tables) {
            if (table.Rows.Count < 2 || table.Rows.Count > rectangles.Length || table.Columns.Count == 0 ||
                PdfTableColumnGeometry.HasHorizontalProgression(table.Columns, static column => column.VisualBounds, false)) continue;
            GetVisualTableBoundaries(table, out double[] rows, out double[] textColumns);
            if (rows.Length != table.Rows.Count + 1 || textColumns.Length < 2) continue;
            var selected = new List<TableStrokeSample>();
            bool ambiguous = false;
            for (int row = 0; row < table.Rows.Count; row++) {
                double centre = (rows[row] + rows[row + 1]) / 2D;
                TableStrokeSample? match = null;
                foreach (TableStrokeSample rectangle in rectangles) {
                    if (rectangle.Left >= centre || rectangle.Right <= centre ||
                        rectangle.Top > textColumns[0] + TableStrokeTolerance ||
                        rectangle.Bottom < textColumns[textColumns.Length - 1] - TableStrokeTolerance) continue;
                    if (match.HasValue && !SameRectangle(match.Value, rectangle)) { ambiguous = true; break; }
                    match = rectangle;
                }
                if (ambiguous || !match.HasValue) { ambiguous = true; break; }
                selected.Add(match.Value);
            }
            if (ambiguous) continue;
            selected.Sort(static (first, second) => first.Left.CompareTo(second.Left));
            TableStrokeSample first = selected[0];
            for (int row = 0; row < selected.Count; row++) {
                TableStrokeSample rectangle = selected[row];
                if (!SameStroke(first, rectangle) || !NearStroke(first.Top, rectangle.Top) ||
                    !NearStroke(first.Bottom, rectangle.Bottom) ||
                    row > 0 && !NearStroke(selected[row - 1].Right, rectangle.Left)) { ambiguous = true; break; }
            }
            if (ambiguous) continue;
            var columns = new List<double> { first.Top, first.Bottom };
            var examined = new List<double>();
            TableStrokeSample[] dividers = samples.Where(line => line.Ordinary &&
                line.Kind == PdfPageVisualPrimitiveKind.Line && line.Bottom - line.Top <= TableStrokeTolerance &&
                SameStroke(first, line) && line.Top > first.Top + TableStrokeTolerance &&
                line.Top < first.Bottom - TableStrokeTolerance).ToArray();
            foreach (TableStrokeSample line in dividers) {
                double coordinate = (line.Top + line.Bottom) / 2D;
                if (examined.Any(value => NearStroke(value, coordinate))) continue;
                examined.Add(coordinate);
                TableStrokeSample[] aligned = dividers.Where(other => NearStroke(other.Top, coordinate)).ToArray();
                if (selected.All(rectangle => aligned.Any(other =>
                    other.Left <= rectangle.Left + TableStrokeTolerance && other.Right >= rectangle.Right - TableStrokeTolerance))) {
                    columns.Add(coordinate);
                }
            }
            result.Add(new PhysicalTableStrokeFrame(selected, columns));
        }
        return result;
    }

    private static bool NearStroke(double first, double second) => Math.Abs(first - second) <= TableStrokeTolerance;
    private static bool SameStroke(TableStrokeSample first, TableStrokeSample second) =>
        first.Color.Equals(second.Color) && Math.Abs(first.StrokeWidth - second.StrokeWidth) <= 0.001D;
    private static bool SameRectangle(TableStrokeSample first, TableStrokeSample second) =>
        NearStroke(first.Left, second.Left) && NearStroke(first.Right, second.Right) &&
        NearStroke(first.Top, second.Top) && NearStroke(first.Bottom, second.Bottom) && SameStroke(first, second);

    private readonly struct TableStrokeSample {
        internal TableStrokeSample(PdfPageVisualPrimitive primitive, bool inferredBorder) {
            Kind = primitive.Kind;
            Left = primitive.X; Top = primitive.Y;
            Right = primitive.X + primitive.Width; Bottom = primitive.Y + primitive.Height;
            StrokeWidth = primitive.StrokeWidth;
            Color = primitive.StrokeColor;
            InferredBorder = inferredBorder;
            Ordinary = primitive.Kind is PdfPageVisualPrimitiveKind.Line or PdfPageVisualPrimitiveKind.Rectangle &&
                primitive.ContentOrderKey?.Depth == 1 && !primitive.HasFillPaint && primitive.HasStrokePaint &&
                Color.HasValue && Color.Value.A == byte.MaxValue && StrokeWidth > 0D && IsFinite(StrokeWidth) &&
                IsFinite(Left) && IsFinite(Top) && IsFinite(Right) && IsFinite(Bottom) &&
                (primitive.StrokeOpacity ?? 1D) == 1D && primitive.ClipPath is null &&
                primitive.StrokeGradient is null && primitive.StrokeRadialGradient is null && primitive.StrokeTilingPattern is null &&
                primitive.StrokeDashStyle == OfficeStrokeDashStyle.Solid && primitive.StrokeDashPattern is not { Array.Count: > 0 };
        }
        internal PdfPageVisualPrimitiveKind Kind { get; }
        internal double Left { get; }
        internal double Top { get; }
        internal double Right { get; }
        internal double Bottom { get; }
        internal double StrokeWidth { get; }
        internal OfficeColor? Color { get; }
        internal bool Ordinary { get; }
        internal bool InferredBorder { get; }
    }

    private sealed class PhysicalTableStrokeFrame {
        private readonly List<TableStrokeSample> _rows;
        private readonly List<double> _columns;
        internal PhysicalTableStrokeFrame(List<TableStrokeSample> rows, List<double> columns) {
            _rows = rows; _columns = columns;
        }
        internal bool Contains(TableStrokeSample sample) =>
            sample.Left >= _rows[0].Left - TableStrokeTolerance &&
            sample.Right <= _rows[_rows.Count - 1].Right + TableStrokeTolerance &&
            sample.Top >= _rows[0].Top - TableStrokeTolerance && sample.Bottom <= _rows[0].Bottom + TableStrokeTolerance;
        internal bool IsBorder(TableStrokeSample sample) {
            if (!sample.Ordinary || !SameStroke(_rows[0], sample)) return false;
            bool LeftOrRight(double value) => _rows.Any(row => NearStroke(row.Left, value) || NearStroke(row.Right, value));
            bool TopOrBottom(double value) => _columns.Any(column => NearStroke(column, value));
            if (sample.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
                return LeftOrRight(sample.Left) && LeftOrRight(sample.Right) && TopOrBottom(sample.Top) && TopOrBottom(sample.Bottom);
            }
            return sample.Right - sample.Left <= TableStrokeTolerance && LeftOrRight((sample.Left + sample.Right) / 2D) ||
                sample.Bottom - sample.Top <= TableStrokeTolerance && TopOrBottom((sample.Top + sample.Bottom) / 2D);
        }
    }
}

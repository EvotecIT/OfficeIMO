using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    /// <summary>
    /// Recognizes root-painted rows on a known opaque backdrop. Each row closes
    /// after its optional fill and ends with full-height interior dividers.
    /// </summary>
    private sealed class PaintedRowRulings {
        private readonly OfficeColor _backdrop;
        private readonly List<Paint> _paint = new();
        private bool _unsupported;

        private PaintedRowRulings(OfficeColor backdrop) => _backdrop = backdrop;

        internal static PaintedRowRulings? Start(PdfPageVisualPrimitive first, double width, double height) =>
            first.ContentOrderKey?.Depth == 1 && first.Kind == PdfPageVisualPrimitiveKind.Rectangle &&
            first.ClipPath is null && !first.HasStrokePaint && IsSolidFill(first) &&
            IsFinite(first.X) && IsFinite(first.Y) && IsFinite(first.Width) && IsFinite(first.Height) &&
            first.X <= 0D && first.Y <= 0D && first.X + first.Width >= width && first.Y + first.Height >= height &&
            first.FillColor is OfficeColor backdrop
                ? new PaintedRowRulings(backdrop) : null;

        internal void Add(PdfPageVisualPrimitive primitive, double width, double height, PdfUnderstandingPageContext context) {
            if (_unsupported) return;
            context.ConsumeWork();
            bool fill = primitive.HasFillPaint;
            bool stroke = primitive.HasStrokePaint;
            if (!fill && !stroke) return;
            if (primitive.ContentOrderKey?.Depth != 1 || fill == stroke ||
                _paint.Count >= TableDetector.MaximumRulingSegments ||
                !IsFinite(primitive.X) || !IsFinite(primitive.Y) ||
                !IsFinite(primitive.Width) || !IsFinite(primitive.Height) ||
                primitive.Width < 0D || primitive.Height < 0D ||
                primitive.FillTilingPattern is not null || primitive.StrokeTilingPattern is not null ||
                primitive.FunctionPaint is not null ||
                !TryGetPaintFloor(primitive, fill, context, out ColorFloor color)) {
                _unsupported = true;
                return;
            }

            double opacity = (fill ? primitive.FillOpacity : primitive.StrokeOpacity) ?? 1D;
            if (!IsFinite(opacity) || opacity < 0D || opacity > 1D) { _unsupported = true; return; }
            if (opacity == 0D) return;
            double halfStroke = stroke ? primitive.StrokeWidth / 2D : 0D;
            if (!IsFinite(halfStroke) || halfStroke < 0D) { _unsupported = true; return; }
            var bounds = new PdfVisualBounds(primitive.X, primitive.Y, primitive.X + primitive.Width, primitive.Y + primitive.Height);
            bool fullyVisible = Contained(bounds.Left - halfStroke, bounds.Top - halfStroke,
                    bounds.Right + halfStroke, bounds.Bottom + halfStroke, 0D, 0D, width, height) &&
                (primitive.ClipPath is null || primitive.ClipPath is { IsExact: true, IsRectangle: true, ContainsTextClipping: false } clip &&
                    Contained(bounds.Left - halfStroke, bounds.Top - halfStroke, bounds.Right + halfStroke, bounds.Bottom + halfStroke,
                        clip.X, clip.Y, clip.X + clip.Width, clip.Y + clip.Height));
            OfficeColor? rule = stroke && fullyVisible && opacity == 1D && halfStroke > 0D &&
                primitive.StrokeColor is OfficeColor ink && ink.A == byte.MaxValue &&
                (ink.R != byte.MaxValue || ink.G != byte.MaxValue || ink.B != byte.MaxValue) &&
                primitive.StrokeGradient is null && primitive.StrokeRadialGradient is null &&
                primitive.StrokeDashStyle == OfficeStrokeDashStyle.Solid && primitive.StrokeDashPattern is not { Array.Count: > 0 }
                    ? ink : null;
            double coveragePad = primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle
                ? halfStroke * 2D
                : halfStroke * Math.Max(2D, primitive.StrokeMiterLimit);
            if (!IsFinite(coveragePad)) { _unsupported = true; return; }
            _paint.Add(new Paint(primitive.Kind, bounds, fill, fill && fullyVisible && IsSolidFill(primitive),
                rule, halfStroke, coveragePad, opacity, color));
        }

        internal void Invalidate() => _unsupported = true;

        internal IReadOnlyList<PdfTableRuling> Build(PdfUnderstandingPageContext context,
            Func<double, double, (double X, double Y)> toReadingPoint) {
            if (_unsupported || _paint.Count == 0 || _paint[_paint.Count - 1].Rule is not OfficeColor rule) {
                return Array.Empty<PdfTableRuling>();
            }
            double halfStroke = _paint[_paint.Count - 1].HalfStroke;
            var rows = new List<Row>();
            int cursor = _paint.Count;
            while (cursor > 0) {
                context.ConsumeWork();
                int end = cursor;
                int start = end;
                while (start > 0 && _paint[start - 1].IsDivider(rule, halfStroke)) {
                    context.ConsumeWork();
                    start--;
                }
                if (start == end || start == 0 || !_paint[start - 1].IsFrame(rule, halfStroke)) break;
                PdfVisualBounds frame = _paint[start - 1].Bounds;
                var columns = new double[end - start];
                double previous = frame.Left;
                for (int index = start; index < end; index++) {
                    context.ConsumeWork();
                    PdfVisualBounds line = _paint[index].Bounds;
                    if (!Near(line.Top, frame.Top) || !Near(line.Bottom, frame.Bottom) ||
                        line.Left <= previous + TableDetector.RulingTolerance ||
                        line.Left >= frame.Right - TableDetector.RulingTolerance) return Array.Empty<PdfTableRuling>();
                    columns[index - start] = previous = line.Left;
                }

                int rowStart = start - 1;
                bool shaded = rowStart > 0 && _paint[rowStart - 1].Fill && SameBox(_paint[rowStart - 1].Bounds, frame);
                if (shaded) {
                    Paint fill = _paint[--rowStart];
                    if (!fill.SolidFill || !fill.Color.LighterThan(rule)) return Array.Empty<PdfTableRuling>();
                }
                rows.Add(new Row(frame, columns));
                cursor = rowStart;
            }
            if (rows.Count < 2) return Array.Empty<PdfTableRuling>();
            rows.Reverse();

            var rulings = new List<PdfTableRuling>();
            for (int start = 0; start < rows.Count;) {
                context.ConsumeWork();
                Row first = rows[start];
                int end = start + 1;
                while (end < rows.Count && Near(rows[end - 1].Bounds.Bottom, rows[end].Bounds.Top)) {
                    context.ConsumeWork();
                    Row next = rows[end];
                    if (!Near(next.Bounds.Left, first.Bounds.Left) || !Near(next.Bounds.Right, first.Bounds.Right) ||
                        next.Columns.Length != first.Columns.Length) return Array.Empty<PdfTableRuling>();
                    for (int column = 0; column < first.Columns.Length; column++) {
                        context.ConsumeWork();
                        if (!Near(next.Columns[column], first.Columns[column])) return Array.Empty<PdfTableRuling>();
                    }
                    end++;
                }
                if (end - start < 2 || end < rows.Count && rows[end].Bounds.Top < rows[end - 1].Bounds.Bottom ||
                    !HasContrastingBackdrop(cursor, first, rows[end - 1].Bounds.Bottom, rule, halfStroke, context)) {
                    return Array.Empty<PdfTableRuling>();
                }

                double left = first.Bounds.Left, right = first.Bounds.Right, top = first.Bounds.Top;
                double bottom = rows[end - 1].Bounds.Bottom;
                long newRulings = (long)(end - start) + 3 + first.Columns.Length;
                if (rulings.Count + newRulings > TableDetector.MaximumRulingSegments) return Array.Empty<PdfTableRuling>();
                Add(left, top, right, top);
                for (int index = start; index < end; index++) Add(left, rows[index].Bounds.Bottom, right, rows[index].Bounds.Bottom);
                Add(left, top, left, bottom);
                foreach (double x in first.Columns) Add(x, top, x, bottom);
                Add(right, top, right, bottom);
                start = end;
            }
            return rulings;

            void Add(double x1, double y1, double x2, double y2) {
                context.ConsumeWork();
                (double X, double Y) first = toReadingPoint(x1, y1), second = toReadingPoint(x2, y2);
                rulings.Add(new PdfTableRuling(first.X, first.Y, second.X, second.Y));
            }
        }

        private bool HasContrastingBackdrop(int prefixCount, Row first, double bottom, OfficeColor rule, double halfStroke,
            PdfUnderstandingPageContext context) {
            PdfVisualBounds topRow = first.Bounds;
            var grid = new PdfVisualBounds(topRow.Left - halfStroke, topRow.Top - halfStroke,
                topRow.Right + halfStroke, bottom + halfStroke);
            var backdrop = new ColorFloor(_backdrop.R, _backdrop.G, _backdrop.B);
            for (int index = 0; index < prefixCount; index++) {
                context.ConsumeWork();
                Paint paint = _paint[index];
                PdfVisualBounds box = paint.Bounds;
                if (paint.Kind == PdfPageVisualPrimitiveKind.Rectangle && !paint.Fill &&
                    Near(box.Left, topRow.Left) && Near(box.Right, topRow.Right) && Near(box.Bottom, topRow.Top)) {
                    // The unmatched prefix might be an earlier row of the same grid.
                    return false;
                }
                double pad = paint.CoveragePad;
                if (box.Right + pad < grid.Left || box.Left - pad > grid.Right ||
                    box.Bottom + pad < grid.Top || box.Top - pad > grid.Bottom) continue;
                if (!paint.Fill && paint.Kind == PdfPageVisualPrimitiveKind.Rectangle &&
                    Contained(grid.Left, grid.Top, grid.Right, grid.Bottom,
                        box.Left + pad, box.Top + pad, box.Right - pad, box.Bottom - pad)) continue;
                backdrop = backdrop.WithPaint(paint.Color, paint.Opacity);
            }
            return backdrop.LighterThan(rule);
        }

        private static bool TryGetPaintFloor(PdfPageVisualPrimitive primitive, bool fill, PdfUnderstandingPageContext context,
            out ColorFloor floor) {
            OfficeColor? color = fill ? primitive.FillColor : primitive.StrokeColor;
            IReadOnlyList<OfficeGradientStop>? stops = fill
                ? primitive.FillGradient?.Stops ?? primitive.FillRadialGradient?.Stops
                : primitive.StrokeGradient?.Stops ?? primitive.StrokeRadialGradient?.Stops;
            if (stops is { Count: > 0 }) {
                byte red = byte.MaxValue, green = byte.MaxValue, blue = byte.MaxValue;
                foreach (OfficeGradientStop stop in stops) {
                    context.ConsumeWork();
                    red = Math.Min(red, stop.Color.R);
                    green = Math.Min(green, stop.Color.G);
                    blue = Math.Min(blue, stop.Color.B);
                }
                floor = new ColorFloor(red, green, blue);
                return true;
            }
            floor = color is OfficeColor solid ? new ColorFloor(solid.R, solid.G, solid.B) : default;
            return color.HasValue;
        }

        private static bool IsSolidFill(PdfPageVisualPrimitive primitive) =>
            primitive.FillColor is { A: byte.MaxValue } && (primitive.FillOpacity ?? 1D) == 1D &&
            primitive.FillGradient is null && primitive.FillRadialGradient is null &&
            primitive.FillTilingPattern is null && primitive.FunctionPaint is null && !primitive.IsSelfColoredShading;

        private static bool Near(double first, double second) => Math.Abs(first - second) <= TableDetector.RulingTolerance;

        private static bool SameBox(PdfVisualBounds first, PdfVisualBounds second) =>
            Near(first.Left, second.Left) && Near(first.Right, second.Right) &&
            Near(first.Top, second.Top) && Near(first.Bottom, second.Bottom);

        private readonly record struct Row(PdfVisualBounds Bounds, double[] Columns);

        private readonly record struct Paint(PdfPageVisualPrimitiveKind Kind, PdfVisualBounds Bounds, bool Fill, bool SolidFill,
            OfficeColor? Rule, double HalfStroke, double CoveragePad, double Opacity, ColorFloor Color) {
            internal bool IsDivider(OfficeColor color, double halfStroke) =>
                Kind == PdfPageVisualPrimitiveKind.Line && Rule == color && HalfStroke == halfStroke &&
                Bounds.Width <= TableDetector.RulingTolerance && Bounds.Height > TableDetector.RulingTolerance;

            internal bool IsFrame(OfficeColor color, double halfStroke) =>
                Kind == PdfPageVisualPrimitiveKind.Rectangle && Rule == color && HalfStroke == halfStroke &&
                Bounds.Width > TableDetector.RulingTolerance && Bounds.Height > TableDetector.RulingTolerance;
        }

        private readonly record struct ColorFloor(double Red, double Green, double Blue) {
            internal bool LighterThan(OfficeColor color) =>
                Red > color.R + 1D || Green > color.G + 1D || Blue > color.B + 1D;

            internal ColorFloor WithPaint(ColorFloor paint, double opacity) =>
                new ColorFloor(Lower(Red, paint.Red, opacity), Lower(Green, paint.Green, opacity), Lower(Blue, paint.Blue, opacity));

            private static double Lower(double background, double paint, double opacity) =>
                Math.Floor(Math.Min(background, background + (paint - background) * opacity));
        }
    }
}

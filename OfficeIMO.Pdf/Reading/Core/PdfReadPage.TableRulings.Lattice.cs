using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    /// <summary>
    /// Collects the borders a reader would see on a page painted the way spreadsheet and word-processor exports
    /// paint tables: cell fills first, then borders as strokes or thin filled rectangles in several colours.
    /// A border counts when its colour stands out from the fill beneath it, and the parts a later opaque fill
    /// or image paints over are removed. These rulings feed lattice recovery only; the strict collector alone
    /// decides whether existing text-table geometry may be refined.
    /// </summary>
    private sealed class LatticeRulingCollector {
        // Borders wider than this read as cell fills rather than rules.
        private const double MaximumRuleThickness = 2D;
        private const double MinimumRuleLength = 1D;
        private const int MaximumFills = 16_384;
        // Colour channels closer than this are treated as the same paint (white borders on a white page).
        private const int MinimumContrast = 16;
        private const int BucketsPerAxis = 32;

        private readonly double _pageWidth;
        private readonly double _pageHeight;
        private readonly Func<double, double, (double X, double Y)> _toReadingPoint;
        private readonly List<Rule> _rules = new();
        private readonly List<Fill> _fills = new();
        private bool _overflow;

        internal LatticeRulingCollector(double pageWidth, double pageHeight, Func<double, double, (double X, double Y)> toReadingPoint) {
            _pageWidth = pageWidth;
            _pageHeight = pageHeight;
            _toReadingPoint = toReadingPoint;
        }

        internal void Add(PdfPageVisualPrimitive primitive, PdfUnderstandingPageContext context) {
            if (_overflow) return;
            context.ConsumeWork();
            var clip = new PdfVisualBounds(0D, 0D, _pageWidth, _pageHeight);
            if (primitive.ClipPath is PdfPageClipPath clipPath) {
                if (clipPath.ContainsTextClipping) return;
                if (!IsFinite(clipPath.X) || !IsFinite(clipPath.Y) || !IsFinite(clipPath.Width) || !IsFinite(clipPath.Height)) return;
                clip = new PdfVisualBounds(Math.Max(clip.Left, clipPath.X), Math.Max(clip.Top, clipPath.Y),
                    Math.Min(clip.Right, clipPath.X + clipPath.Width), Math.Min(clip.Bottom, clipPath.Y + clipPath.Height));
                if (clip.Right <= clip.Left || clip.Bottom <= clip.Top) return;
            }
            double order = primitive.PaintOrder;
            if (primitive.HasFillPaint) AddFill(primitive, clip, order, context);
            if (primitive.HasStrokePaint) AddStroke(primitive, clip, order, context);
        }

        private void AddFill(PdfPageVisualPrimitive primitive, PdfVisualBounds clip, double order, PdfUnderstandingPageContext context) {
            double opacity = primitive.FillOpacity ?? 1D;
            if (!IsFinite(opacity) || opacity <= 0D) return;
            bool solid = primitive.FillColor is { } fillColor && fillColor.A == byte.MaxValue &&
                primitive.FillGradient is null && primitive.FillRadialGradient is null &&
                primitive.FillTilingPattern is null && primitive.FunctionPaint is null && !primitive.IsSelfColoredShading;
            OfficeColor? color = solid ? primitive.FillColor : null;
            bool opaque = opacity >= 1D && (color is null || color.Value.A == byte.MaxValue);
            if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
                AddFilledBox(primitive.X, primitive.Y, primitive.X + primitive.Width, primitive.Y + primitive.Height,
                    clip, order, color, opaque);
                return;
            }
            if (primitive.Kind != PdfPageVisualPrimitiveKind.Path) return;
            // Producers fill many borders as one path of box subpaths. Each axis-aligned box counts on its own;
            // curved or slanted subpaths are artwork and are skipped.
            OfficePoint first = default, previous = default;
            int corners = 0;
            double left = 0D, top = 0D, right = 0D, bottom = 0D;
            bool open = false, box = false;
            foreach (OfficePathCommand command in primitive.PathCommands) {
                context.ConsumeWork();
                switch (command.Kind) {
                    case OfficePathCommandKind.MoveTo:
                        CloseBox();
                        first = previous = command.Point;
                        left = right = first.X; top = bottom = first.Y;
                        corners = 1;
                        open = box = true;
                        break;
                    case OfficePathCommandKind.LineTo when open:
                        if (!Orthogonal(previous, command.Point)) box = false;
                        previous = command.Point;
                        left = Math.Min(left, previous.X); right = Math.Max(right, previous.X);
                        top = Math.Min(top, previous.Y); bottom = Math.Max(bottom, previous.Y);
                        corners++;
                        break;
                    case OfficePathCommandKind.Close when open:
                        CloseBox();
                        break;
                    default:
                        box = false;
                        break;
                }
                if (_overflow) return;
            }
            CloseBox();

            void CloseBox() {
                if (open && box && Orthogonal(previous, first) && corners is >= 4 and <= 5) {
                    AddFilledBox(left, top, right, bottom, clip, order, color, opaque);
                }
                open = box = false;
            }
        }

        private void AddFilledBox(double x1, double y1, double x2, double y2, PdfVisualBounds clip, double order,
            OfficeColor? color, bool opaque) {
            double left = Math.Max(Math.Min(x1, x2), clip.Left), right = Math.Min(Math.Max(x1, x2), clip.Right);
            double top = Math.Max(Math.Min(y1, y2), clip.Top), bottom = Math.Min(Math.Max(y1, y2), clip.Bottom);
            if (!IsFinite(left) || !IsFinite(right) || !IsFinite(top) || !IsFinite(bottom) || right <= left || bottom <= top) return;
            double width = right - left, height = bottom - top;
            double thickness = Math.Min(width, height), length = Math.Max(width, height);
            if (thickness <= MaximumRuleThickness && length >= Math.Max(MinimumRuleLength, thickness * 2D) && color.HasValue) {
                // Producers leave the joint square out of the bars on either side of it; extending each bar
                // by its thickness reconnects them.
                if (width >= height) AddRule(left - height, top + height / 2D, right + height, top + height / 2D, height, order, color.Value, opaque);
                else AddRule(left + width / 2D, top - width, left + width / 2D, bottom + width, width, order, color.Value, opaque);
                return;
            }
            if (thickness <= MaximumRuleThickness) return;
            if (_fills.Count >= MaximumFills) { _overflow = true; return; }
            (double X, double Y) a = _toReadingPoint(left, top), b = _toReadingPoint(right, bottom);
            _fills.Add(new Fill(Math.Min(a.X, b.X), Math.Min(a.Y, b.Y), Math.Max(a.X, b.X), Math.Max(a.Y, b.Y), order, color, opaque));
        }

        private void AddStroke(PdfPageVisualPrimitive primitive, PdfVisualBounds clip, double order, PdfUnderstandingPageContext context) {
            double opacity = primitive.StrokeOpacity ?? 1D;
            double width = primitive.StrokeWidth;
            if (!IsFinite(opacity) || opacity <= 0D || !IsFinite(width) || width <= 0D || width > MaximumRuleThickness ||
                primitive.StrokeColor is not { } color || color.A == 0 ||
                primitive.StrokeGradient is not null || primitive.StrokeRadialGradient is not null ||
                primitive.StrokeTilingPattern is not null) return;
            bool opaque = opacity >= 1D && color.A == byte.MaxValue;
            if (primitive.Kind == PdfPageVisualPrimitiveKind.Line) {
                AddClippedRule(primitive.X1, primitive.Y1, primitive.X2, primitive.Y2, width, clip, order, color, opaque);
                return;
            }
            if (primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle) {
                double right = primitive.X + primitive.Width, bottom = primitive.Y + primitive.Height;
                AddClippedRule(primitive.X, primitive.Y, right, primitive.Y, width, clip, order, color, opaque);
                AddClippedRule(right, primitive.Y, right, bottom, width, clip, order, color, opaque);
                AddClippedRule(right, bottom, primitive.X, bottom, width, clip, order, color, opaque);
                AddClippedRule(primitive.X, bottom, primitive.X, primitive.Y, width, clip, order, color, opaque);
                return;
            }
            OfficePoint? previous = null, start = null;
            foreach (OfficePathCommand command in primitive.PathCommands) {
                context.ConsumeWork();
                if (command.Kind == OfficePathCommandKind.MoveTo) previous = start = command.Point;
                else if (command.Kind == OfficePathCommandKind.LineTo && previous.HasValue) {
                    AddClippedRule(previous.Value.X, previous.Value.Y, command.Point.X, command.Point.Y, width, clip, order, color, opaque);
                    previous = command.Point;
                } else if (command.Kind == OfficePathCommandKind.Close && previous.HasValue && start.HasValue) {
                    AddClippedRule(previous.Value.X, previous.Value.Y, start.Value.X, start.Value.Y, width, clip, order, color, opaque);
                    previous = start;
                } else {
                    // Curves are artwork; keep the straight pieces already read and skip the rest of the path.
                    return;
                }
            }
        }

        private void AddClippedRule(double x1, double y1, double x2, double y2, double thickness, PdfVisualBounds clip,
            double order, OfficeColor color, bool opaque) {
            if (!IsFinite(x1) || !IsFinite(y1) || !IsFinite(x2) || !IsFinite(y2)) return;
            if (Math.Abs(y1 - y2) <= TableDetector.RulingTolerance) {
                if (y1 < clip.Top || y1 > clip.Bottom) return;
                double from = Math.Max(Math.Min(x1, x2), clip.Left), to = Math.Min(Math.Max(x1, x2), clip.Right);
                if (to <= from) return;
                x1 = from; x2 = to; y2 = y1;
            } else if (Math.Abs(x1 - x2) <= TableDetector.RulingTolerance) {
                if (x1 < clip.Left || x1 > clip.Right) return;
                double from = Math.Max(Math.Min(y1, y2), clip.Top), to = Math.Min(Math.Max(y1, y2), clip.Bottom);
                if (to <= from) return;
                y1 = from; y2 = to; x2 = x1;
            } else {
                return;
            }
            AddRule(x1, y1, x2, y2, thickness, order, color, opaque);
        }

        private void AddRule(double x1, double y1, double x2, double y2, double thickness, double order, OfficeColor color, bool opaque) {
            if (_rules.Count >= TableDetector.MaximumRulingSegments) { _overflow = true; return; }
            (double X, double Y) a = _toReadingPoint(x1, y1), b = _toReadingPoint(x2, y2);
            if (!IsFinite(a.X) || !IsFinite(a.Y) || !IsFinite(b.X) || !IsFinite(b.Y)) return;
            bool horizontal = Math.Abs(a.Y - b.Y) <= Math.Abs(a.X - b.X);
            double lo = horizontal ? Math.Min(a.X, b.X) : Math.Min(a.Y, b.Y);
            double hi = horizontal ? Math.Max(a.X, b.X) : Math.Max(a.Y, b.Y);
            if (hi - lo < MinimumRuleLength) return;
            double position = horizontal ? (a.Y + b.Y) / 2D : (a.X + b.X) / 2D;
            _rules.Add(new Rule(horizontal, position, lo, hi, thickness / 2D, order, color, opaque));
        }

        internal IReadOnlyList<PdfTableRuling> Build(PdfUnderstandingPageContext context) {
            if (_overflow || _rules.Count < 4) return Array.Empty<PdfTableRuling>();
            List<Fill> fills = _fills;
            foreach (PdfImagePlacement image in context.ImagePlacements) {
                context.ConsumeWork();
                if (!IsFinite(image.X) || !IsFinite(image.Y) || !IsFinite(image.Width) || !IsFinite(image.Height)) continue;
                (double X, double Y) a = ToFrame(image.X, image.Y), b = ToFrame(image.X + image.Width, image.Y + image.Height);
                fills.Add(new Fill(Math.Min(a.X, b.X), Math.Min(a.Y, b.Y), Math.Max(a.X, b.X), Math.Max(a.Y, b.Y),
                    image.PaintOrder, null, image.Opacity >= 1D));
            }
            var index = new FillIndex(fills, context);
            // Rules by orientation and whole-point position, so a rule finds the later rules painted over it.
            var byLine = new Dictionary<long, List<int>>();
            for (int ruleIndex = 0; ruleIndex < _rules.Count; ruleIndex++) {
                long key = LineKey(_rules[ruleIndex].Horizontal, _rules[ruleIndex].Position);
                if (!byLine.TryGetValue(key, out List<int>? list)) byLine.Add(key, list = new List<int>());
                list.Add(ruleIndex);
            }
            var result = new List<PdfTableRuling>();
            var cuts = new List<(double From, double To)>();
            var candidates = new List<int>();
            foreach (Rule rule in _rules) {
                context.ConsumeWork();
                double middle = (rule.Lo + rule.Hi) / 2D;
                double backdropOrder = double.NegativeInfinity;
                OfficeColor? backdrop = OfficeColor.White;
                cuts.Clear();
                double left = rule.Horizontal ? rule.Lo : rule.Position - rule.Half;
                double right = rule.Horizontal ? rule.Hi : rule.Position + rule.Half;
                double bottom = rule.Horizontal ? rule.Position - rule.Half : rule.Lo;
                double top = rule.Horizontal ? rule.Position + rule.Half : rule.Hi;
                index.Query(left, bottom, right, top, candidates);
                foreach (int candidate in candidates) {
                    context.ConsumeWork();
                    Fill fill = fills[candidate];
                    if (!fill.Opaque) continue;
                    if (fill.Order < rule.Order) {
                        double x = rule.Horizontal ? middle : rule.Position, y = rule.Horizontal ? rule.Position : middle;
                        if (fill.Order > backdropOrder && x >= fill.Left && x <= fill.Right && y >= fill.Bottom && y <= fill.Top) {
                            backdropOrder = fill.Order;
                            backdrop = fill.Color;
                        }
                    } else if (fill.Order > rule.Order) {
                        // Paint over the whole thickness hides that stretch of the border.
                        double near = rule.Horizontal ? fill.Bottom : fill.Left, far = rule.Horizontal ? fill.Top : fill.Right;
                        if (near <= rule.Position - rule.Half + 0.05D && far >= rule.Position + rule.Half - 0.05D) {
                            cuts.Add(rule.Horizontal ? (fill.Left, fill.Right) : (fill.Bottom, fill.Top));
                        }
                    }
                }
                // A later opaque stroke or bar along the same line covering the whole thickness paints this
                // stretch over, whatever its own colour: a white stroke over a black divider erases it.
                for (long key = LineKey(rule.Horizontal, rule.Position) - 2; key <= LineKey(rule.Horizontal, rule.Position) + 2; key++) {
                    if (!byLine.TryGetValue(key, out List<int>? sameLine)) continue;
                    foreach (int other in sameLine) {
                        context.ConsumeWork();
                        Rule cover = _rules[other];
                        if (!cover.Opaque || cover.Order <= rule.Order || cover.Horizontal != rule.Horizontal ||
                            Math.Abs(cover.Position - rule.Position) + rule.Half > cover.Half + 0.05D) continue;
                        cuts.Add((cover.Lo, cover.Hi));
                    }
                }
                if (backdrop is { } under && !Contrasts(rule.Color, under)) continue;
                AddVisible(rule, cuts, result, context);
                if (result.Count > TableDetector.MaximumRulingSegments) return Array.Empty<PdfTableRuling>();
            }
            return result;

            (double X, double Y) ToFrame(double x, double y) =>
                context.ReadingFrame is null ? (x, y) : context.ReadingFrame.ToFrame(x, y);
        }

        private static void AddVisible(Rule rule, List<(double From, double To)> cuts, List<PdfTableRuling> result,
            PdfUnderstandingPageContext context) {
            if (cuts.Count > 1) cuts.Sort(static (a, b) => a.From.CompareTo(b.From));
            double from = rule.Lo;
            foreach ((double cutFrom, double cutTo) in cuts) {
                context.ConsumeWork();
                if (cutTo <= from) continue;
                if (cutFrom >= rule.Hi) break;
                if (cutFrom - from >= MinimumRuleLength) Emit(from, cutFrom);
                from = Math.Max(from, cutTo);
            }
            if (rule.Hi - from >= MinimumRuleLength) Emit(from, rule.Hi);

            void Emit(double lo, double hi) => result.Add(rule.Horizontal
                ? new PdfTableRuling(lo, rule.Position, hi, rule.Position)
                : new PdfTableRuling(rule.Position, lo, rule.Position, hi));
        }

        private static long LineKey(bool horizontal, double position) =>
            (horizontal ? 1L << 40 : 0L) + (long)Math.Floor(position);

        private static bool Contrasts(OfficeColor ink, OfficeColor backdrop) =>
            Math.Abs(ink.R - backdrop.R) >= MinimumContrast || Math.Abs(ink.G - backdrop.G) >= MinimumContrast ||
            Math.Abs(ink.B - backdrop.B) >= MinimumContrast;

        private static bool Orthogonal(OfficePoint a, OfficePoint b) =>
            Math.Abs(a.X - b.X) <= TableDetector.RulingTolerance || Math.Abs(a.Y - b.Y) <= TableDetector.RulingTolerance;

        private readonly record struct Rule(bool Horizontal, double Position, double Lo, double Hi, double Half, double Order, OfficeColor Color, bool Opaque);

        private readonly record struct Fill(double Left, double Bottom, double Right, double Top, double Order, OfficeColor? Color, bool Opaque);

        /// <summary>A coarse uniform grid over the fills so each border checks only the paint near it.</summary>
        private sealed class FillIndex {
            private readonly List<int>?[] _buckets = new List<int>?[BucketsPerAxis * BucketsPerAxis];
            private readonly int[] _stamps;
            private readonly double _left, _bottom, _cellWidth, _cellHeight;
            private int _stamp;

            internal FillIndex(List<Fill> fills, PdfUnderstandingPageContext context) {
                _stamps = new int[fills.Count];
                if (fills.Count == 0) { _cellWidth = _cellHeight = 1D; return; }
                double left = double.MaxValue, bottom = double.MaxValue, right = double.MinValue, top = double.MinValue;
                foreach (Fill fill in fills) {
                    left = Math.Min(left, fill.Left); bottom = Math.Min(bottom, fill.Bottom);
                    right = Math.Max(right, fill.Right); top = Math.Max(top, fill.Top);
                }
                _left = left; _bottom = bottom;
                _cellWidth = Math.Max((right - left) / BucketsPerAxis, 1D);
                _cellHeight = Math.Max((top - bottom) / BucketsPerAxis, 1D);
                for (int fillIndex = 0; fillIndex < fills.Count; fillIndex++) {
                    Fill fill = fills[fillIndex];
                    (int x0, int y0, int x1, int y1) = Cells(fill.Left, fill.Bottom, fill.Right, fill.Top);
                    for (int y = y0; y <= y1; y++) {
                        for (int x = x0; x <= x1; x++) {
                            context.ConsumeWork();
                            (_buckets[y * BucketsPerAxis + x] ??= new List<int>()).Add(fillIndex);
                        }
                    }
                }
            }

            internal void Query(double left, double bottom, double right, double top, List<int> found) {
                found.Clear();
                if (_stamps.Length == 0) return;
                _stamp++;
                (int x0, int y0, int x1, int y1) = Cells(left, bottom, right, top);
                for (int y = y0; y <= y1; y++) {
                    for (int x = x0; x <= x1; x++) {
                        List<int>? bucket = _buckets[y * BucketsPerAxis + x];
                        if (bucket is null) continue;
                        foreach (int fill in bucket) {
                            if (_stamps[fill] == _stamp) continue;
                            _stamps[fill] = _stamp;
                            found.Add(fill);
                        }
                    }
                }
            }

            private (int X0, int Y0, int X1, int Y1) Cells(double left, double bottom, double right, double top) =>
                (Cell(left, _left, _cellWidth), Cell(bottom, _bottom, _cellHeight),
                    Cell(right, _left, _cellWidth), Cell(top, _bottom, _cellHeight));

            private static int Cell(double value, double origin, double size) {
                double cell = Math.Floor((value - origin) / size);
                return cell < 0D ? 0 : cell >= BucketsPerAxis ? BucketsPerAxis - 1 : (int)cell;
            }
        }
    }
}

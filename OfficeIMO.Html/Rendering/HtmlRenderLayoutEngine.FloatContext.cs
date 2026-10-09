namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static bool EstablishesIndependentFloatContext(HtmlRenderBoxStyle style) =>
        style.Display == "flow-root" || style.Display == "inline-block"
        || style.Display == "flex" || style.Display == "grid" || style.Display == "table"
        || (style.OverflowX != "visible" && style.OverflowX != "clip")
        || (style.OverflowY != "visible" && style.OverflowY != "clip")
        || IsVerticalWritingMode(style.WritingMode);

    private static HtmlRenderFlowBlock CreateFloatClearanceBlock(double width, double height) =>
        new HtmlRenderFlowBlock(width, height, Array.Empty<HtmlRenderVisual>(),
            HtmlPageBreakTarget.None, HtmlPageBreakTarget.None, false, "float-clearance");

    private HtmlRenderBoxStyle PlaceIndependentBlockBesideFloats(HtmlRenderBoxStyle style, double width,
        InlineFloatContext context, ref double y) {
        if (!context.HasFloats || !EstablishesIndependentFloatContext(style)) return style;
        double height = Math.Max(style.LineHeight, style.ExplicitHeight ?? 0D);
        InlineFloatBand band = context.ResolveUsableBand(ref y, height);
        while (ResolveBoxWidth(Math.Max(1D, band.Width - style.MarginLeft - style.MarginRight), style)
            + style.MarginLeft + style.MarginRight > band.Width + 0.0001D) {
            double next = context.NextBottomAfter(y);
            if (next <= y + 0.0001D) break;
            y = next;
            band = context.ResolveUsableBand(ref y, height);
        }
        if (band.Left <= 0D && band.Right >= width) return style;
        HtmlRenderBoxStyle placed = style.Clone();
        placed.MarginLeft += band.Left;
        placed.MarginRight += width - band.Right;
        return placed;
    }

    private sealed class InlineFloatContext {
        private readonly double _width;
        private readonly List<InlineFloatPlacement> _placements;
        private readonly double _originX, _originY;
        private readonly IReadOnlyList<HtmlFloatExclusion> _inheritedFloats;
        private readonly List<HtmlFloatExclusion> _originatingPageExclusions;
        internal IReadOnlyList<HtmlFloatExclusion> OriginatingPageExclusions => _originatingPageExclusions
            .Select(item => item.Shift(-_originX, -_originY)).ToArray();

        internal void ReserveOriginatingPageExclusion(InlineFloatPlacement placement, double start, double end) {
            if (end <= start + 0.0001D) return;
            _originatingPageExclusions.Add(new HtmlFloatExclusion(
                placement.X + _originX, start + _originY, placement.Width, end - start, placement.Run.FloatSide));
        }

        internal double FirstLineIndent;
        internal bool IndentAtRight;

        internal InlineFloatContext(double width, IReadOnlyList<HtmlFloatExclusion>? inheritedFloats = null)
            : this(width, new List<InlineFloatPlacement>(), 0D, 0D,
                inheritedFloats ?? Array.Empty<HtmlFloatExclusion>(), new List<HtmlFloatExclusion>()) { }

        private InlineFloatContext(double width, List<InlineFloatPlacement> placements, double originX, double originY, IReadOnlyList<HtmlFloatExclusion> inheritedFloats,
            List<HtmlFloatExclusion> originatingPageExclusions) {
            _width = Math.Max(1D, width);
            _placements = placements;
            _originX = originX;
            _originY = originY;
            _inheritedFloats = inheritedFloats;
            _originatingPageExclusions = originatingPageExclusions;
        }

        internal bool HasFloats => _placements.Count > 0 || _inheritedFloats.Count > 0 || _originatingPageExclusions.Count > 0;
        internal InlineFloatContext At(double width, double x, double y) =>
            new InlineFloatContext(width, _placements, _originX + x, _originY + y, _inheritedFloats, _originatingPageExclusions);

        private void Register(InlineFloatPlacement placement) => _placements.Add(
            new InlineFloatPlacement(placement.Run, placement.X + _originX, placement.Y + _originY, placement.Width, placement.Height));

        internal double Bottom => _placements.Count == 0 ? 0D : Math.Max(0D, _placements.Max(item => item.Bottom) - _originY);

        internal InlineFloatPlacement Place(HtmlInlineRun run, double requestedY) {
            HtmlRenderFlowBlock block = run.FloatingBlock!;
            double boxWidth = Math.Min(_width, Math.Max(0.01D, block.Width));
            double boxHeight = Math.Max(0.01D, block.Height);
            double y = Math.Max(requestedY, Clearance(run.ClearSide));
            while (true) {
                InlineFloatBand band = ResolveBand(y, boxHeight);
                if (boxWidth <= band.Width + 0.0001D) {
                    double x = run.FloatSide == "right" ? band.Right - boxWidth : band.Left;
                    var placement = new InlineFloatPlacement(run, x, y, boxWidth, boxHeight);
                    Register(placement);
                    return placement;
                }
                double next = NextBottomAfter(y);
                if (next <= y + 0.0001D) {
                    double x = run.FloatSide == "right" ? Math.Max(0D, _width - boxWidth) : 0D;
                    var placement = new InlineFloatPlacement(run, x, y, boxWidth, boxHeight);
                    Register(placement);
                    return placement;
                }
                y = next;
            }
        }

        internal InlineFloatBand ResolveUsableBand(ref double y, double height) {
            InlineFloatBand band = ResolveBand(y, height);
            while (band.Width <= 0.01D) {
                double next = NextBottomAfter(y);
                if (next <= y + 0.0001D) break;
                y = next;
                band = ResolveBand(y, height);
            }
            return band;
        }

        internal InlineFloatBand ResolveBand(double y, double height) {
            double left = 0D;
            double right = _width;
            double bottom = y + Math.Max(0.01D, height);
            foreach (HtmlFloatExclusion exclusion in _inheritedFloats.Concat(_originatingPageExclusions)) {
                if (exclusion.Y >= bottom + _originY - 0.0001D || exclusion.Bottom <= y + _originY + 0.0001D) continue;
                if (exclusion.Right <= _originX || exclusion.X >= _originX + _width) continue;
                if (exclusion.Side == "right") right = Math.Max(0D, Math.Min(right, exclusion.X - _originX));
                else left = Math.Min(_width, Math.Max(left, exclusion.Right - _originX));
            }
            foreach (InlineFloatPlacement placement in _placements) {
                if (placement.Y >= bottom + _originY - 0.0001D || placement.Bottom <= y + _originY + 0.0001D) continue;
                if (placement.Right <= _originX || placement.X >= _originX + _width) continue;
                if (placement.Run.FloatSide == "right") right = Math.Max(0D, Math.Min(right, placement.X - _originX));
                else left = Math.Min(_width, Math.Max(left, placement.Right - _originX));
            }
            return new InlineFloatBand(left, Math.Max(left, right));
        }

        internal double NextBottomAfter(double y) => _placements.Select(item => item.Bottom)
            .Concat(_inheritedFloats.Select(item => item.Bottom))
            .Concat(_originatingPageExclusions.Select(item => item.Bottom))
            .Where(bottom => bottom > y + _originY + 0.0001D)
            .Select(bottom => bottom - _originY)
            .DefaultIfEmpty(y)
            .Min();

        internal double Clearance(string clearSide) {
            if (clearSide == "none") return 0D;
            return _placements
                .Where(item => clearSide == "both" || item.Run.FloatSide == clearSide)
                .Select(item => item.Bottom - _originY)
                .Concat(_inheritedFloats.Where(item => clearSide == "both" || item.Side == clearSide).Select(item => item.Bottom - _originY))
                .Concat(_originatingPageExclusions.Where(item => clearSide == "both" || item.Side == clearSide).Select(item => item.Bottom - _originY))
                .DefaultIfEmpty(0D)
                .Max();
        }
    }

}

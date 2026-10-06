namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Paint only this edge and its rounded corners. Clipping a whole rectangle
    // would leave perpendicular stubs at square internal cell boundaries.
    private static void DrawRoundedPairedCellBorderSide(StringBuilder sb, PdfCellBorderSide border, RoundedRectSide side,
        double x, double y, double w, double h, double radius, bool tl, bool tr, bool br, bool bl,
        PdfCellBorderSide? startSide, PdfCellBorderSide? endSide, bool artifact) {
        PaintTrack(-GetCellBorderPairOutset(border), -GetCellBorderPairOutset(startSide), -GetCellBorderPairOutset(endSide));
        PaintTrack(GetCellBorderPairInset(border), GetCellBorderPairInset(startSide), GetCellBorderPairInset(endSide));

        void PaintTrack(double requestedInset, double startInset, double endInset) {
            double inset = Math.Min(requestedInset, Math.Min(w, h) / 2D);
            if (w - inset * 2D <= 0D || h - inset * 2D <= 0D) return;
            double r = Math.Max(0D, Math.Min(radius, Math.Min(w, h) / 2D) - inset);
            double c = r * 0.5522847498307936D;
            AppendArtifactBegin(sb, artifact);
            var content = new ContentStreamBuilder(sb).SaveState().StrokeColor(border.Color!.Value).LineWidth(border.Width);
            ApplyStrokeDashStyle(content, border.DashStyle, border.Width, hasExplicitLineCap: false);
            switch (side) {
                case RoundedRectSide.Top: {
                    double top = y + h - inset;
                    double left = x + (tl ? inset : startInset), right = x + w - (tr ? inset : endInset);
                    content.MoveTo(left, top - (tl ? r : 0D));
                    if (tl && r > 0D) content.CubicTo(left, top - r + c, left + r - c, top, left + r, top);
                    content.LineTo(right - (tr ? r : 0D), top);
                    if (tr && r > 0D) content.CubicTo(right - r + c, top, right, top - r + c, right, top - r);
                    break;
                }
                case RoundedRectSide.Right: {
                    double right = x + w - inset;
                    double top = y + h - (tr ? inset : startInset), bottom = y + (br ? inset : endInset);
                    content.MoveTo(right - (tr ? r : 0D), top);
                    if (tr && r > 0D) content.CubicTo(right - r + c, top, right, top - r + c, right, top - r);
                    content.LineTo(right, bottom + (br ? r : 0D));
                    if (br && r > 0D) content.CubicTo(right, bottom + r - c, right - r + c, bottom, right - r, bottom);
                    break;
                }
                case RoundedRectSide.Bottom: {
                    double bottom = y + inset;
                    double left = x + (bl ? inset : startInset), right = x + w - (br ? inset : endInset);
                    content.MoveTo(left, bottom + (bl ? r : 0D));
                    if (bl && r > 0D) content.CubicTo(left, bottom + r - c, left + r - c, bottom, left + r, bottom);
                    content.LineTo(right - (br ? r : 0D), bottom);
                    if (br && r > 0D) content.CubicTo(right - r + c, bottom, right, bottom + r - c, right, bottom + r);
                    break;
                }
                default: {
                    double left = x + inset;
                    double top = y + h - (tl ? inset : startInset), bottom = y + (bl ? inset : endInset);
                    content.MoveTo(left + (tl ? r : 0D), top);
                    if (tl && r > 0D) content.CubicTo(left + r - c, top, left, top - r + c, left, top - r);
                    content.LineTo(left, bottom + (bl ? r : 0D));
                    if (bl && r > 0D) content.CubicTo(left, bottom + r - c, left + r - c, bottom, left + r, bottom);
                    break;
                }
            }
            content.StrokePath().RestoreState();
            AppendArtifactEnd(sb, artifact);
        }
    }
}

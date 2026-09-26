namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly List<(LayoutResult.Page? Page, double Left, double Right, double Top, double Bottom, bool AllowOverlap)> floatingTables = new();

        private double PositionTableX(PdfTablePosition position, double tableWidth) {
            double left = position.HorizontalAnchor == PdfTableAnchor.Page ? 0 : currentOpts.MarginLeft;
            double available = position.HorizontalAnchor == PdfTableAnchor.Page ? currentOpts.PageWidth : width;
            PdfAlign horizontal = position.HorizontalAlignment;
            if (position.MirrorHorizontalOnEvenPages && (pages.Count + 1) % 2 == 0)
                horizontal = horizontal == PdfAlign.Left ? PdfAlign.Right : horizontal == PdfAlign.Right ? PdfAlign.Left : horizontal;
            double alignment = horizontal == PdfAlign.Right ? available - tableWidth
                : horizontal == PdfAlign.Center ? (available - tableWidth) / 2 : 0;
            return left + alignment + position.HorizontalOffset;
        }

        private double PositionTableY(PdfTablePosition position, double tableHeight) {
            double top = position.VerticalAnchor == PdfTableAnchor.Page ? currentOpts.PageHeight
                : position.VerticalAnchor == PdfTableAnchor.Margin ? GetCurrentFramePageStartY() : y;
            double bottom = position.VerticalAnchor == PdfTableAnchor.Page ? 0 : currentOpts.MarginBottom;
            double alignment = position.VerticalAlignment == PdfTableVerticalAlignment.Bottom ? top - bottom - tableHeight
                : position.VerticalAlignment == PdfTableVerticalAlignment.Center ? (top - bottom - tableHeight) / 2 : 0;
            return top - alignment - position.VerticalOffset;
        }

        private void AvoidFloatingTable(PdfTablePosition position, double left, double tableWidth, double height) {
            bool moved;
            do {
                moved = false;
                foreach (var region in floatingTables) {
                    if ((position.AllowOverlap && region.AllowOverlap) || !ReferenceEquals(region.Page, currentPage) || left + tableWidth <= region.Left || left >= region.Right ||
                        y - height >= region.Top || y <= region.Bottom) continue;
                    y = region.Bottom - position.DistanceTop;
                    moved = true;
                }
            } while (moved);
        }

        private void ReserveFloatingTable(PdfTablePosition position, double left, double top, double tableWidth, double height) {
            floatingTables.RemoveAll(region => !ReferenceEquals(region.Page, currentPage));
            floatingTables.Add((currentPage, left - position.DistanceLeft,
                left + tableWidth + position.DistanceRight, top + position.DistanceTop,
                top - height - position.DistanceBottom, position.AllowOverlap));
        }

        private bool HasFloatingTables => floatingTables.Any(region => ReferenceEquals(region.Page, currentPage));

        private void AvoidFloatingBlock(double height) {
            bool moved;
            do {
                moved = false;
                foreach (var region in floatingTables) {
                    if (!ReferenceEquals(region.Page, currentPage) || y <= region.Bottom || y - height >= region.Top ||
                        region.Right <= currentOpts.MarginLeft || region.Left >= currentOpts.MarginLeft + width) continue;
                    y = region.Bottom;
                    moved = true;
                }
            } while (moved);
        }

        // Choose the widest connected text interval; a line that cannot fit moves below the obstruction.
        private (double X, double Width, double Gap) GetFloatingTextFrame(double left, double lineWidth, double top, double height, double minimumWidth = 0) {
            double originalTop = top;
            for (int attempt = 0; attempt <= floatingTables.Count; attempt++) {
                var intervals = new List<(double Left, double Right)> { (left, left + lineWidth) };
                double nextBottom = top;
                foreach (var region in floatingTables) {
                    if (!ReferenceEquals(region.Page, currentPage) || top - height >= region.Top || top <= region.Bottom) continue;
                    if (region.Right <= left || region.Left >= left + lineWidth) continue;
                    nextBottom = Math.Min(nextBottom, region.Bottom);
                    var remaining = new List<(double Left, double Right)>();
                    foreach (var interval in intervals) {
                        if (region.Right <= interval.Left || region.Left >= interval.Right) { remaining.Add(interval); continue; }
                        if (region.Left > interval.Left) remaining.Add((interval.Left, Math.Min(interval.Right, region.Left)));
                        if (region.Right < interval.Right) remaining.Add((Math.Max(interval.Left, region.Right), interval.Right));
                    }
                    intervals = remaining;
                }
                var best = intervals.OrderByDescending(interval => interval.Right - interval.Left).FirstOrDefault();
                if (best.Right - best.Left >= Math.Min(lineWidth, Math.Max(minimumWidth, Math.Max(24, currentOpts.DefaultFontSize * 2))))
                    return (best.Left, best.Right - best.Left, originalTop - top);
                if (nextBottom >= top) break;
                top = nextBottom;
            }
            return (left, lineWidth, originalTop - top);
        }
    }
}

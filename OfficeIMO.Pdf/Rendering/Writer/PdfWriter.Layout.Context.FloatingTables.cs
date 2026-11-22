namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly List<(LayoutResult.Page? Page, double Left, double Right, double Top, double Bottom, bool AllowOverlap)> floatingTables = new();

        private sealed class FloatingFlowCapture {
            public List<PdfLayoutRegion> PaintedRegions { get; } = new();
            public HashSet<int> FlowPages { get; } = new();
        }
        private readonly Stack<FloatingFlowCapture> activeFloatingFlowCaptures = new();

        private void RecordFlowPlacement(double topY) {
            foreach (var capture in activeFloatingFlowCaptures) capture.FlowPages.Add(pages.Count + 1);
            ResolveFloatingBookmarks(topY);
        }

        private readonly List<(LayoutResult.Page Page, PageNamedDestination Destination)> pendingFloatingBookmarks = new();

        private void QueueFloatingBookmark(string name) {
            EnsurePage();
            var destination = new PageNamedDestination { Name = name, Y = y, AttachToNextStructure = true };
            currentPage!.NamedDestinations.Add(destination);
            pendingFloatingBookmarks.Add((currentPage, destination));
        }

        private void ResolveFloatingBookmarks(double topY) {
            if (pendingFloatingBookmarks.Count == 0) return;
            Dictionary<LayoutResult.Page, HashSet<PageNamedDestination>>? movedByPage = null;
            foreach (var pending in pendingFloatingBookmarks) {
                pending.Destination.Y = topY;
                if (!ReferenceEquals(pending.Page, currentPage)) {
                    movedByPage ??= new Dictionary<LayoutResult.Page, HashSet<PageNamedDestination>>();
                    if (!movedByPage.TryGetValue(pending.Page, out HashSet<PageNamedDestination>? moved)) {
                        moved = new HashSet<PageNamedDestination>();
                        movedByPage.Add(pending.Page, moved);
                    }
                    moved.Add(pending.Destination);
                    currentPage!.NamedDestinations.Add(pending.Destination);
                }
            }
            if (movedByPage != null) {
                foreach (var moved in movedByPage)
                    moved.Key.NamedDestinations.RemoveAll(moved.Value.Contains);
            }
            pendingFloatingBookmarks.Clear();
        }

        // A page group continued on this page anchors to its own side margins.
        // A header or footer is laid out in its own context; its physical page decides parity.
        private int PhysicalPageNumberForPosition => PicturePage;

        private PdfCreateOptions PageAnchorOptions => activeContainerScopes.Count > 0 ? activeContainerScopes[0].PageOptions
            : continuedPageAnchorOptions ?? currentPage?.Options ?? currentOpts;

        private const double WordBorderlessFloatingTableRightEdge = 0.75D;
        private const double WordMinimumFloatingTableTextDistance = 0.5D;

        /// <summary>The paint a floating table carries beyond its grid on the left and right, as Word measures it.</summary>
        private static (double Left, double Right) GetFloatingTableOutsideEdges(PdfTablePosition position, PdfTableStyle? style) {
            if (!position.UsesWordOutsideEdges) return (0D, 0D);
            double left = 0D, right = 0D;
            if (style?.CellBorders is { Count: > 0 } borders) {
                int lastColumn = -1;
                PdfCellBorder? last = null;
                foreach (var entry in borders) {
                    if (entry.Key.Row != 0) continue;
                    if (entry.Key.Column == 0 && entry.Value.Left)
                        left = RenderableBorderWidth(ResolveCellBorderSide(entry.Value.LeftBorderSnapshot, entry.Value));
                    if (entry.Key.Column > lastColumn) { lastColumn = entry.Key.Column; last = entry.Value; }
                }
                if (last is { Right: true }) right = RenderableBorderWidth(ResolveCellBorderSide(last.RightBorderSnapshot, last));
            }
            return (left / 2D, right > 0D ? right / 2D : WordBorderlessFloatingTableRightEdge);

            static double RenderableBorderWidth(PdfCellBorderSide? side) => IsRenderableCellBorderSide(side) ? side!.Width.ToPoints() : 0D;
        }

        private static (double Left, double Right) GetFloatingTableTextDistances(PdfTablePosition position) {
            double left = position.DistanceLeft.ToPoints(), right = position.DistanceRight.ToPoints();
            return position.UsesWordOutsideEdges
                ? (Math.Max(left, WordMinimumFloatingTableTextDistance), Math.Max(right, WordMinimumFloatingTableTextDistance))
                : (left, right);
        }

        /// <summary>Returns the grid's left edge; alignment and offsets place the table's outside edges.</summary>
        private double PositionTableX(PdfTablePosition position, double tableWidth, PdfTableStyle? style) {
            var outside = GetFloatingTableOutsideEdges(position, style);
            return PositionTableX(position, tableWidth + outside.Left + outside.Right) + outside.Left;
        }

        private double PositionTableX(PdfTablePosition position, double tableWidth) {
            double left = position.HorizontalAnchor == PdfTableAnchor.Page ? 0
                : position.HorizontalAnchor == PdfTableAnchor.Margin ? PageAnchorOptions.MarginLeft.ToPoints() : currentOpts.MarginLeft.ToPoints();
            double available = position.HorizontalAnchor == PdfTableAnchor.Page ? PageAnchorOptions.PageWidth.ToPoints()
                : position.HorizontalAnchor == PdfTableAnchor.Margin ? PageAnchorOptions.PageWidth.ToPoints() - PageAnchorOptions.MarginLeft.ToPoints() - PageAnchorOptions.MarginRight.ToPoints() : width;
            OfficeAlignment horizontal = position.HorizontalAlignment;
            if (position.MirrorHorizontalOnEvenPages && PhysicalPageNumberForPosition % 2 == 0)
                horizontal = horizontal == OfficeAlignment.Left ? OfficeAlignment.Right : horizontal == OfficeAlignment.Right ? OfficeAlignment.Left : horizontal;
            if (position.AlignsInPageMargin && position.HorizontalAnchor == PdfTableAnchor.Page && horizontal != OfficeAlignment.Center) {
                double pageWidth = PageAnchorOptions.PageWidth.ToPoints();
                double marginLeft = PageAnchorOptions.MarginLeft.ToPoints();
                double marginRight = PageAnchorOptions.MarginRight.ToPoints();
                double inMargin = horizontal == OfficeAlignment.Right
                    ? Math.Min(pageWidth - marginRight, pageWidth - tableWidth)
                    : Math.Max(0D, marginLeft - tableWidth);
                return inMargin + position.HorizontalOffset.ToPoints();
            }
            double alignment = horizontal == OfficeAlignment.Right ? available - tableWidth
                : horizontal == OfficeAlignment.Center ? (available - tableWidth) / 2 : 0;
            return left + alignment + position.HorizontalOffset.ToPoints();
        }

        private double PositionTableY(PdfTablePosition position, double tableHeight) {
            double top = position.VerticalAnchor == PdfTableAnchor.Page ? currentOpts.PageHeight.ToPoints()
                : position.VerticalAnchor == PdfTableAnchor.Margin ? PageAnchorOptions.PageHeight.ToPoints() - PageAnchorOptions.MarginTop.ToPoints() : y;
            double bottom = position.VerticalAnchor == PdfTableAnchor.Page ? 0
                : position.VerticalAnchor == PdfTableAnchor.Margin ? PageAnchorOptions.MarginBottom.ToPoints() : currentOpts.MarginBottom.ToPoints();
            PdfTableVerticalAlignment vertical = position.VerticalAlignment;
            bool evenPage = PhysicalPageNumberForPosition % 2 == 0;
            if (vertical == PdfTableVerticalAlignment.Inside) vertical = evenPage ? PdfTableVerticalAlignment.Bottom : PdfTableVerticalAlignment.Top;
            else if (vertical == PdfTableVerticalAlignment.Outside) vertical = evenPage ? PdfTableVerticalAlignment.Top : PdfTableVerticalAlignment.Bottom;
            double alignment = vertical == PdfTableVerticalAlignment.Bottom ? top - bottom - tableHeight
                : vertical == PdfTableVerticalAlignment.Center ? (top - bottom - tableHeight) / 2 : 0;
            return top - alignment - position.VerticalOffset.ToPoints();
        }

        private void AvoidFloatingTable(PdfTablePosition position, double left, double tableWidth, double height) {
            bool moved;
            do {
                moved = false;
                foreach (var region in floatingTables) {
                    if ((position.AllowOverlap && region.AllowOverlap) || !ReferenceEquals(region.Page, currentPage) || left + tableWidth <= region.Left || left >= region.Right ||
                        y - height >= region.Top || y <= region.Bottom) continue;
                    y = region.Bottom - position.DistanceTop.ToPoints();
                    moved = true;
                }
            } while (moved);
        }

        private void ReserveFloatingTable(PdfTablePosition position, double left, double top, double tableWidth, double height, PdfTableStyle? style) {
            foreach (var capture in activeFloatingFlowCaptures)
                capture.PaintedRegions.Add(new PdfLayoutRegion(pages.Count + 1, left, top - height, tableWidth, height));
            floatingTables.RemoveAll(region => !ReferenceEquals(region.Page, currentPage));
            // Word lets other floats overlap a frame that text runs underneath.
            if (!position.WrapsText) return;
            var outside = GetFloatingTableOutsideEdges(position, style);
            var distance = GetFloatingTableTextDistances(position);
            floatingTables.Add((currentPage, left - outside.Left - distance.Left,
                left + tableWidth + outside.Right + distance.Right, top + position.DistanceTop.ToPoints(),
                top - height - position.DistanceBottom.ToPoints(), position.AllowOverlap));
        }

        private bool HasFloatingTables => floatingTables.Any(region => ReferenceEquals(region.Page, currentPage));

        private void AvoidFloatingBlock(double height) {
            bool moved;
            do {
                moved = false;
                foreach (var region in floatingTables) {
                    if (!ReferenceEquals(region.Page, currentPage) || y <= region.Bottom || y - height >= region.Top ||
                        region.Right <= currentOpts.MarginLeft.ToPoints() || region.Left >= currentOpts.MarginLeft.ToPoints() + width) continue;
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
                    intervals = SubtractTextExclusion(intervals.Select(interval =>
                            new RichTextInterval(interval.Left, interval.Right - interval.Left)).ToArray(), region.Left, region.Right)
                        .Select(interval => (Left: interval.X, Right: interval.X + interval.Width)).ToList();
                }
                var best = intervals.OrderByDescending(interval => interval.Right - interval.Left).FirstOrDefault();
                if (best.Right - best.Left >= Math.Min(lineWidth, Math.Max(minimumWidth, Math.Max(24, currentOpts.DefaultFontSize.ToPoints() * 2))))
                    return (best.Left, best.Right - best.Left, originalTop - top);
                if (nextBottom >= top) break;
                top = nextBottom;
            }
            return (left, lineWidth, originalTop - top);
        }
    }
}

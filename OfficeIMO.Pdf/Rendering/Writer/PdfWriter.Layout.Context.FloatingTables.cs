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

        private readonly record struct FloatingTableOutsideEdges(double Placement, double Left, double Right);

        private readonly System.Runtime.CompilerServices.ConditionalWeakTable<TableBlock, Tuple<PdfTableStyle, FloatingTableOutsideEdges>> floatingTableEdges = new();

        /// <summary>
        /// The paint a floating table carries beyond its grid, as Word measures it. The first row's
        /// outer left border places the grid; text clears the widest outer border on each side
        /// across all rows. A cell without its own borders takes the table's uniform grid border,
        /// and a table with no right border anywhere still reaches 0.75pt past its grid.
        /// </summary>
        private FloatingTableOutsideEdges GetFloatingTableOutsideEdges(PdfTablePosition position, TableBlock table, PdfTableStyle style) {
            if (!position.UsesWordOutsideEdges) return default;
            if (floatingTableEdges.TryGetValue(table, out var cached) && ReferenceEquals(cached.Item1, style)) return cached.Item2;
            int columns = GetTableColumnCount(table);
            double placement = 0D, left = 0D, right = 0D;
            bool firstRow = true;
            for (int row = 0; row < table.Rows.Count; row++) {
                List<TableCellLayout> cells = GetTableCellLayouts(table, row, columns);
                if (cells.Count == 0) continue;
                TableCellLayout first = cells[0], last = cells[0];
                foreach (TableCellLayout cell in cells) {
                    if (cell.Column < first.Column) first = cell;
                    if (cell.Column + cell.ColumnSpan > last.Column + last.ColumnSpan) last = cell;
                }
                double rowLeft = first.Column == 0 ? OuterHalfWidth(row, first.Column, rightSide: false) : 0D;
                double rowRight = last.Column + last.ColumnSpan >= columns ? OuterHalfWidth(row, last.Column, rightSide: true) : 0D;
                if (firstRow && first.Column == 0) { placement = rowLeft; firstRow = false; }
                left = Math.Max(left, rowLeft);
                right = Math.Max(right, rowRight);
            }
            var edges = new FloatingTableOutsideEdges(placement, left, right > 0D ? right : WordBorderlessFloatingTableRightEdge);
            floatingTableEdges.Remove(table);
            floatingTableEdges.Add(table, Tuple.Create(style, edges));
            return edges;

            double OuterHalfWidth(int row, int column, bool rightSide) {
                if (style.CellBorders == null || !style.CellBorders.TryGetValue((row, column), out PdfCellBorder? border))
                    return style.BorderColor.HasValue ? style.BorderWidth.ToPoints() / 2D : 0D;
                PdfCellBorderSide? side = rightSide
                    ? border.Right ? ResolveCellBorderSide(border.RightBorderSnapshot, border) : null
                    : border.Left ? ResolveCellBorderSide(border.LeftBorderSnapshot, border) : null;
                return IsRenderableCellBorderSide(side) ? side!.Width.ToPoints() / 2D : 0D;
            }
        }

        private static (double Left, double Right) GetFloatingTableTextDistances(PdfTablePosition position) {
            double left = position.DistanceLeft.ToPoints(), right = position.DistanceRight.ToPoints();
            return position.UsesWordOutsideEdges
                ? (Math.Max(left, WordMinimumFloatingTableTextDistance), Math.Max(right, WordMinimumFloatingTableTextDistance))
                : (left, right);
        }

        /// <summary>Returns the grid's left edge; alignment and offsets place the table's outside edges.</summary>
        private double PositionTableX(PdfTablePosition position, double tableWidth, TableBlock table, PdfTableStyle style) {
            var outside = GetFloatingTableOutsideEdges(position, table, style);
            return PositionTableX(position, tableWidth + outside.Placement + outside.Right) + outside.Placement;
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
                double offset = position.HorizontalOffset.ToPoints();
                // A left frame that would leave the page stops at its edge, after its offset;
                // the text beside it then clears the frame where it stands.
                return horizontal == OfficeAlignment.Right
                    ? Math.Min(pageWidth - marginRight, pageWidth - tableWidth) + offset
                    : Math.Max(0D, marginLeft - tableWidth + offset);
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

        private void ReserveFloatingTable(PdfTablePosition position, double left, double top, double tableWidth, double height, TableBlock table, PdfTableStyle style) {
            foreach (var capture in activeFloatingFlowCaptures)
                capture.PaintedRegions.Add(new PdfLayoutRegion(pages.Count + 1, left, top - height, tableWidth, height));
            floatingTables.RemoveAll(region => !ReferenceEquals(region.Page, currentPage));
            // Word lets other floats overlap a frame that text runs underneath.
            if (!position.WrapsText) return;
            var outside = GetFloatingTableOutsideEdges(position, table, style);
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
    }
}

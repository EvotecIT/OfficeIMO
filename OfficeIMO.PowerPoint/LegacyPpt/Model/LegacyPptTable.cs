using OfficeIMO.Drawing.Binary;
using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents a native binary PowerPoint table decoded from an OfficeArt shape group.</summary>
    public sealed class LegacyPptTable {
        private const ushort TablePropertiesId = 0x039F;

        private LegacyPptTable(IReadOnlyList<int> columnBoundaries,
            IReadOnlyList<int> rowBoundaries,
            IReadOnlyList<LegacyPptTableCell> cells,
            byte? styleFlags,
            bool hasExplicitGridLines) {
            ColumnWidths = new ReadOnlyCollection<int>(CreateSizes(columnBoundaries));
            RowHeights = new ReadOnlyCollection<int>(CreateSizes(rowBoundaries));
            Cells = new ReadOnlyCollection<LegacyPptTableCell>(cells.ToArray());
            FirstRow = HasStyleFlag(styleFlags, 0);
            LastRow = HasStyleFlag(styleFlags, 1);
            FirstColumn = HasStyleFlag(styleFlags, 2);
            LastColumn = HasStyleFlag(styleFlags, 3);
            BandedRows = HasStyleFlag(styleFlags, 4);
            BandedColumns = HasStyleFlag(styleFlags, 5);
            HasExplicitGridLines = hasExplicitGridLines;
        }

        /// <summary>Gets the number of table rows.</summary>
        public int Rows => RowHeights.Count;

        /// <summary>Gets the number of table columns.</summary>
        public int Columns => ColumnWidths.Count;

        /// <summary>Gets column widths in the table group's master-unit coordinate system.</summary>
        public IReadOnlyList<int> ColumnWidths { get; }

        /// <summary>Gets row heights in the table group's master-unit coordinate system.</summary>
        public IReadOnlyList<int> RowHeights { get; }

        /// <summary>Gets native table cells in row-major drawing order.</summary>
        public IReadOnlyList<LegacyPptTableCell> Cells { get; }

        /// <summary>Gets whether first-row styling was retained by the binary producer.</summary>
        public bool FirstRow { get; }

        /// <summary>Gets whether last-row styling was retained by the binary producer.</summary>
        public bool LastRow { get; }

        /// <summary>Gets whether first-column styling was retained by the binary producer.</summary>
        public bool FirstColumn { get; }

        /// <summary>Gets whether last-column styling was retained by the binary producer.</summary>
        public bool LastColumn { get; }

        /// <summary>Gets whether alternating row styling was retained by the binary producer.</summary>
        public bool BandedRows { get; }

        /// <summary>Gets whether alternating column styling was retained by the binary producer.</summary>
        public bool BandedColumns { get; }

        internal bool HasExplicitGridLines { get; }

        internal static LegacyPptTable? TryCreate(OfficeArtShapeStyle style,
            IReadOnlyList<LegacyPptShape> children,
            byte? styleFlags = null) {
            if (style == null) throw new ArgumentNullException(nameof(style));
            if (children == null) throw new ArgumentNullException(nameof(children));
            OfficeArtProperty? tableMarker = style.Properties.LastOrDefault(
                property => property.PropertyId == TablePropertiesId
                    && !property.IsComplex && !property.IsBlipId);
            if (tableMarker == null || (tableMarker.Value & 1U) == 0U) return null;

            LegacyPptShape[] candidates = children.Where(IsCellCandidate).ToArray();
            if (candidates.Length == 0) return null;
            int[] horizontal = candidates.SelectMany(shape => new[] {
                    shape.Bounds.Left,
                    checked(shape.Bounds.Left + shape.Bounds.Width)
                })
                .Distinct().OrderBy(value => value).ToArray();
            int[] vertical = candidates.SelectMany(shape => new[] {
                    shape.Bounds.Top,
                    checked(shape.Bounds.Top + shape.Bounds.Height)
                })
                .Distinct().OrderBy(value => value).ToArray();
            if (horizontal.Length < 2 || vertical.Length < 2) return null;

            IReadOnlyDictionary<(int Left, int Top, int Width, int Height),
                LegacyPptShape> gridLines = children
                .Where(IsGridLineCandidate)
                .GroupBy(shape => (shape.Bounds.Left, shape.Bounds.Top,
                    shape.Bounds.Width, shape.Bounds.Height))
                .ToDictionary(group => group.Key, group => group.Last());
            var cells = new List<LegacyPptTableCell>(candidates.Length);
            foreach (LegacyPptShape candidate in candidates) {
                int left = Array.BinarySearch(horizontal, candidate.Bounds.Left);
                int right = Array.BinarySearch(horizontal,
                    checked(candidate.Bounds.Left + candidate.Bounds.Width));
                int top = Array.BinarySearch(vertical, candidate.Bounds.Top);
                int bottom = Array.BinarySearch(vertical,
                    checked(candidate.Bounds.Top + candidate.Bounds.Height));
                if (left < 0 || right <= left || top < 0 || bottom <= top) {
                    return null;
                }
                int cellLeft = horizontal[left];
                int cellRight = horizontal[right];
                int cellTop = vertical[top];
                int cellBottom = vertical[bottom];
                cells.Add(new LegacyPptTableCell(top, left,
                    bottom - top, right - left, candidate,
                    FindBorder(gridLines, cellLeft, cellTop, 0,
                        cellBottom - cellTop),
                    FindBorder(gridLines, cellLeft, cellTop,
                        cellRight - cellLeft, 0),
                    FindBorder(gridLines, cellRight, cellTop, 0,
                        cellBottom - cellTop),
                    FindBorder(gridLines, cellLeft, cellBottom,
                        cellRight - cellLeft, 0)));
            }
            if (cells.GroupBy(cell => (cell.Row, cell.Column)).Any(group => group.Count() > 1)) {
                return null;
            }
            return new LegacyPptTable(horizontal, vertical, cells
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToArray(), styleFlags, gridLines.Count > 0);
        }

        private static bool IsCellCandidate(LegacyPptShape shape) =>
            shape.Bounds.Width > 0 && shape.Bounds.Height > 0
            && shape.Kind is LegacyPptShapeKind.TextBox
                or LegacyPptShapeKind.Rectangle
                or LegacyPptShapeKind.AutoShape;

        private static bool IsGridLineCandidate(LegacyPptShape shape) =>
            shape.Kind == LegacyPptShapeKind.Line
            && (shape.Bounds.Width == 0 || shape.Bounds.Height == 0);

        private static LegacyPptTableBorder? FindBorder(
            IReadOnlyDictionary<(int Left, int Top, int Width, int Height),
                LegacyPptShape> gridLines, int left, int top,
            int width, int height) {
            if (!gridLines.TryGetValue((left, top, width, height),
                    out LegacyPptShape? line)) {
                return null;
            }
            if (line.Style.LineEnabled == false
                || line.Style.LineWidthEmus is <= 0) {
                return new LegacyPptTableBorder(isVisible: false,
                    color: line.LineColor ?? "000000", widthPoints: 0D);
            }
            return new LegacyPptTableBorder(isVisible: true,
                color: line.LineColor ?? "000000",
                widthPoints: line.Style.LineWidthEmus.HasValue
                    ? line.Style.LineWidthEmus.Value / 12700D
                    : 0.75D);
        }

        private static bool HasStyleFlag(byte? flags, int bit) =>
            flags.HasValue && (flags.Value & (1 << bit)) != 0;

        private static int[] CreateSizes(IReadOnlyList<int> boundaries) {
            var sizes = new int[boundaries.Count - 1];
            for (int index = 0; index < sizes.Length; index++) {
                sizes[index] = checked(boundaries[index + 1] - boundaries[index]);
            }
            return sizes;
        }
    }

    /// <summary>Represents one native binary PowerPoint table cell.</summary>
    public sealed class LegacyPptTableCell {
        internal LegacyPptTableCell(int row, int column, int rowSpan,
            int columnSpan, LegacyPptShape sourceShape,
            LegacyPptTableBorder? leftBorder,
            LegacyPptTableBorder? topBorder,
            LegacyPptTableBorder? rightBorder,
            LegacyPptTableBorder? bottomBorder) {
            Row = row;
            Column = column;
            RowSpan = rowSpan;
            ColumnSpan = columnSpan;
            SourceShape = sourceShape ?? throw new ArgumentNullException(nameof(sourceShape));
            LeftBorder = leftBorder;
            TopBorder = topBorder;
            RightBorder = rightBorder;
            BottomBorder = bottomBorder;
        }

        /// <summary>Gets the zero-based starting row.</summary>
        public int Row { get; }

        /// <summary>Gets the zero-based starting column.</summary>
        public int Column { get; }

        /// <summary>Gets the number of rows occupied by the cell.</summary>
        public int RowSpan { get; }

        /// <summary>Gets the number of columns occupied by the cell.</summary>
        public int ColumnSpan { get; }

        /// <summary>Gets the source OfficeArt cell shape and its text/style metadata.</summary>
        public LegacyPptShape SourceShape { get; }

        /// <summary>Gets the resolved left border, including an explicit invisible border.</summary>
        public LegacyPptTableBorder? LeftBorder { get; }

        /// <summary>Gets the resolved top border, including an explicit invisible border.</summary>
        public LegacyPptTableBorder? TopBorder { get; }

        /// <summary>Gets the resolved right border, including an explicit invisible border.</summary>
        public LegacyPptTableBorder? RightBorder { get; }

        /// <summary>Gets the resolved bottom border, including an explicit invisible border.</summary>
        public LegacyPptTableBorder? BottomBorder { get; }
    }

    /// <summary>Describes one native binary PowerPoint table border.</summary>
    public readonly struct LegacyPptTableBorder {
        internal LegacyPptTableBorder(bool isVisible, string color,
            double widthPoints) {
            IsVisible = isVisible;
            Color = color;
            WidthPoints = widthPoints;
        }

        /// <summary>Gets whether the border is explicitly visible.</summary>
        public bool IsVisible { get; }

        /// <summary>Gets the resolved border color as RRGGBB.</summary>
        public string Color { get; }

        /// <summary>Gets the border width in points.</summary>
        public double WidthPoints { get; }
    }
}

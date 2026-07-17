using OfficeIMO.Drawing.Binary;
using System.Collections.ObjectModel;

namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents a native binary PowerPoint table decoded from an OfficeArt shape group.</summary>
    public sealed class LegacyPptTable {
        private const ushort TablePropertiesId = 0x039F;

        private LegacyPptTable(IReadOnlyList<int> columnBoundaries,
            IReadOnlyList<int> rowBoundaries,
            IReadOnlyList<LegacyPptTableCell> cells) {
            ColumnWidths = new ReadOnlyCollection<int>(CreateSizes(columnBoundaries));
            RowHeights = new ReadOnlyCollection<int>(CreateSizes(rowBoundaries));
            Cells = new ReadOnlyCollection<LegacyPptTableCell>(cells.ToArray());
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

        internal static LegacyPptTable? TryCreate(OfficeArtShapeStyle style,
            IReadOnlyList<LegacyPptShape> children) {
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
                cells.Add(new LegacyPptTableCell(top, left,
                    bottom - top, right - left, candidate));
            }
            if (cells.GroupBy(cell => (cell.Row, cell.Column)).Any(group => group.Count() > 1)) {
                return null;
            }
            return new LegacyPptTable(horizontal, vertical, cells
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToArray());
        }

        private static bool IsCellCandidate(LegacyPptShape shape) =>
            shape.Bounds.Width > 0 && shape.Bounds.Height > 0
            && shape.Kind is LegacyPptShapeKind.TextBox
                or LegacyPptShapeKind.Rectangle
                or LegacyPptShapeKind.AutoShape;

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
            int columnSpan, LegacyPptShape sourceShape) {
            Row = row;
            Column = column;
            RowSpan = rowSpan;
            ColumnSpan = columnSpan;
            SourceShape = sourceShape ?? throw new ArgumentNullException(nameof(sourceShape));
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
    }
}

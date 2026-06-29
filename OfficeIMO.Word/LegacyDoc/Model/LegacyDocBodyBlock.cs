namespace OfficeIMO.Word.LegacyDoc.Model {
    internal enum LegacyDocTableCellHorizontalMerge {
        None,
        Restart,
        Continue
    }

    internal enum LegacyDocTableCellVerticalMerge {
        None,
        Restart,
        Continue
    }

    internal enum LegacyDocTableCellVerticalAlignment {
        Top,
        Center,
        Bottom
    }

    internal readonly struct LegacyDocTableCellMargins : IEquatable<LegacyDocTableCellMargins> {
        internal LegacyDocTableCellMargins(int? topTwips, int? rightTwips, int? bottomTwips, int? leftTwips) {
            TopTwips = topTwips;
            RightTwips = rightTwips;
            BottomTwips = bottomTwips;
            LeftTwips = leftTwips;
        }

        internal int? TopTwips { get; }

        internal int? RightTwips { get; }

        internal int? BottomTwips { get; }

        internal int? LeftTwips { get; }

        internal bool HasAny => TopTwips != null || RightTwips != null || BottomTwips != null || LeftTwips != null;

        internal LegacyDocTableCellMargins Merge(LegacyDocTableCellMargins margins) {
            return new LegacyDocTableCellMargins(
                margins.TopTwips ?? TopTwips,
                margins.RightTwips ?? RightTwips,
                margins.BottomTwips ?? BottomTwips,
                margins.LeftTwips ?? LeftTwips);
        }

        public bool Equals(LegacyDocTableCellMargins other) {
            return TopTwips == other.TopTwips
                && RightTwips == other.RightTwips
                && BottomTwips == other.BottomTwips
                && LeftTwips == other.LeftTwips;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocTableCellMargins other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + TopTwips.GetHashCode();
            hash = (hash * 31) + RightTwips.GetHashCode();
            hash = (hash * 31) + BottomTwips.GetHashCode();
            hash = (hash * 31) + LeftTwips.GetHashCode();
            return hash;
        }
    }

    internal abstract class LegacyDocBodyBlock {
        private protected LegacyDocBodyBlock() {
        }
    }

    internal sealed class LegacyDocParagraphBlock : LegacyDocBodyBlock {
        internal LegacyDocParagraphBlock(IReadOnlyList<LegacyDocTextRun> runs, LegacyDocParagraphFormat format) {
            Runs = runs;
            Format = format;
        }

        internal IReadOnlyList<LegacyDocTextRun> Runs { get; }

        internal LegacyDocParagraphFormat Format { get; }
    }

    internal sealed class LegacyDocSectionBreakBlock : LegacyDocBodyBlock {
        internal LegacyDocSectionBreakBlock(LegacyDocSectionFormat format) {
            Format = format;
        }

        internal LegacyDocSectionFormat Format { get; }
    }

    internal sealed class LegacyDocTableBlock : LegacyDocBodyBlock {
        internal LegacyDocTableBlock(IReadOnlyList<LegacyDocTableRow> rows) {
            Rows = rows;
        }

        internal IReadOnlyList<LegacyDocTableRow> Rows { get; }
    }

    internal sealed class LegacyDocTableRow {
        internal LegacyDocTableRow(
            IReadOnlyList<LegacyDocTableCell> cells,
            IReadOnlyList<int>? cellWidthsTwips = null,
            int? rowHeightTwips = null,
            bool rowHeightIsExact = false,
            bool? rowCantSplit = null,
            bool? rowIsHeader = null,
            IReadOnlyList<LegacyDocTableCellHorizontalMerge>? cellHorizontalMerges = null,
            IReadOnlyList<LegacyDocTableCellVerticalMerge>? cellVerticalMerges = null,
            IReadOnlyList<LegacyDocTableCellVerticalAlignment>? cellVerticalAlignments = null,
            IReadOnlyList<bool>? cellFitTexts = null,
            IReadOnlyList<bool>? cellNoWraps = null,
            IReadOnlyList<LegacyDocTableCellMargins>? cellMargins = null) {
            Cells = cells;
            CellWidthsTwips = cellWidthsTwips == null || cellWidthsTwips.Count == 0
                ? Array.Empty<int>()
                : cellWidthsTwips.ToArray();
            RowHeightTwips = rowHeightTwips;
            RowHeightIsExact = rowHeightIsExact;
            RowCantSplit = rowCantSplit;
            RowIsHeader = rowIsHeader;
            CellHorizontalMerges = cellHorizontalMerges == null || cellHorizontalMerges.Count == 0
                ? Array.Empty<LegacyDocTableCellHorizontalMerge>()
                : cellHorizontalMerges.ToArray();
            CellVerticalMerges = cellVerticalMerges == null || cellVerticalMerges.Count == 0
                ? Array.Empty<LegacyDocTableCellVerticalMerge>()
                : cellVerticalMerges.ToArray();
            CellVerticalAlignments = cellVerticalAlignments == null || cellVerticalAlignments.Count == 0
                ? Array.Empty<LegacyDocTableCellVerticalAlignment>()
                : cellVerticalAlignments.ToArray();
            CellFitTexts = cellFitTexts == null || cellFitTexts.Count == 0
                ? Array.Empty<bool>()
                : cellFitTexts.ToArray();
            CellNoWraps = cellNoWraps == null || cellNoWraps.Count == 0
                ? Array.Empty<bool>()
                : cellNoWraps.ToArray();
            CellMargins = cellMargins == null || cellMargins.Count == 0
                ? Array.Empty<LegacyDocTableCellMargins>()
                : cellMargins.ToArray();
        }

        internal IReadOnlyList<LegacyDocTableCell> Cells { get; }

        internal IReadOnlyList<int> CellWidthsTwips { get; }

        internal int? RowHeightTwips { get; }

        internal bool RowHeightIsExact { get; }

        internal bool? RowCantSplit { get; }

        internal bool? RowIsHeader { get; }

        internal IReadOnlyList<LegacyDocTableCellHorizontalMerge> CellHorizontalMerges { get; }

        internal IReadOnlyList<LegacyDocTableCellVerticalMerge> CellVerticalMerges { get; }

        internal IReadOnlyList<LegacyDocTableCellVerticalAlignment> CellVerticalAlignments { get; }

        internal IReadOnlyList<bool> CellFitTexts { get; }

        internal IReadOnlyList<bool> CellNoWraps { get; }

        internal IReadOnlyList<LegacyDocTableCellMargins> CellMargins { get; }
    }

    internal sealed class LegacyDocTableCell {
        internal LegacyDocTableCell(IReadOnlyList<LegacyDocTextRun> runs, LegacyDocParagraphFormat format) {
            Runs = runs;
            Format = format;
            Text = string.Concat(runs.Select(run => run.Text));
        }

        internal string Text { get; }

        internal IReadOnlyList<LegacyDocTextRun> Runs { get; }

        internal LegacyDocParagraphFormat Format { get; }
    }
}

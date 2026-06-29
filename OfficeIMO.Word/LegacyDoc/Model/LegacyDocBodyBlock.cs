namespace OfficeIMO.Word.LegacyDoc.Model {
    internal enum LegacyDocTableCellHorizontalMerge {
        None,
        Restart,
        Continue
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
            IReadOnlyList<LegacyDocTableCellHorizontalMerge>? cellHorizontalMerges = null) {
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
        }

        internal IReadOnlyList<LegacyDocTableCell> Cells { get; }

        internal IReadOnlyList<int> CellWidthsTwips { get; }

        internal int? RowHeightTwips { get; }

        internal bool RowHeightIsExact { get; }

        internal bool? RowCantSplit { get; }

        internal bool? RowIsHeader { get; }

        internal IReadOnlyList<LegacyDocTableCellHorizontalMerge> CellHorizontalMerges { get; }
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

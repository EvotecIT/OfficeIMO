namespace OfficeIMO.Word.LegacyDoc.Model {
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

    internal sealed class LegacyDocTableBlock : LegacyDocBodyBlock {
        internal LegacyDocTableBlock(IReadOnlyList<LegacyDocTableRow> rows) {
            Rows = rows;
        }

        internal IReadOnlyList<LegacyDocTableRow> Rows { get; }
    }

    internal sealed class LegacyDocTableRow {
        internal LegacyDocTableRow(IReadOnlyList<LegacyDocTableCell> cells) {
            Cells = cells;
        }

        internal IReadOnlyList<LegacyDocTableCell> Cells { get; }
    }

    internal sealed class LegacyDocTableCell {
        internal LegacyDocTableCell(IReadOnlyList<LegacyDocTextRun> runs) {
            Runs = runs;
            Text = string.Concat(runs.Select(run => run.Text));
        }

        internal string Text { get; }

        internal IReadOnlyList<LegacyDocTextRun> Runs { get; }
    }
}

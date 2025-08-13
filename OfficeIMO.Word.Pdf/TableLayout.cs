using System.Collections.Generic;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Pdf {
    internal sealed class TableLayout {
        internal TableLayout(List<IReadOnlyList<WordTableCell>> rows, float[] columnWidths) {
            Rows = rows;
            ColumnWidths = columnWidths;
        }

        internal List<IReadOnlyList<WordTableCell>> Rows { get; }

        internal float[] ColumnWidths { get; }
    }
}


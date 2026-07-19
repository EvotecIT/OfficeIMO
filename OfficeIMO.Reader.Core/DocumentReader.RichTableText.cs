using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    internal static string BuildRichTableText(ReaderTable table) {
        IEnumerable<IReadOnlyList<string>> rows = table.Columns.Count == 0
            ? table.Rows
            : new[] { table.Columns }.Concat(table.Rows);
        return string.Join(
            Environment.NewLine,
            rows.Select(static row => string.Join(" | ", row)));
    }
}

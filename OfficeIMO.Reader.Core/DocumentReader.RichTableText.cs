using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    internal static string BuildRichTableText(ReaderTable table) =>
        BuildRichTableText(table, CancellationToken.None);

    internal static string BuildRichTableText(ReaderTable table, CancellationToken cancellationToken) {
        IEnumerable<IReadOnlyList<string>> rows = table.Columns.Count == 0
            ? table.Rows
            : new[] { table.Columns }.Concat(table.Rows);
        var builder = new StringBuilder();
        bool first = true;
        foreach (IReadOnlyList<string> row in rows) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!first) builder.Append(Environment.NewLine);
            builder.Append(string.Join(" | ", row));
            first = false;
        }
        return builder.ToString();
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal static partial class WordRunFormatting {
    private static readonly AsyncLocal<RunPaintConversion?> RunPaintConversions = new();

    /// <summary>Caches read-only table positions for one conversion; nesting restores the caller's scope.</summary>
    internal static IDisposable BeginConversionScope(CancellationToken cancellationToken = default) {
        RunPaintConversion? previous = RunPaintConversions.Value;
        RunPaintConversions.Value = new RunPaintConversion(cancellationToken);
        return new RunPaintScope(previous);
    }

    private sealed class RunPaintConversion {
        internal readonly CancellationToken CancellationToken;
        internal readonly Dictionary<Table, TableRunPaintState> Tables = new();
        internal RunPaintConversion(CancellationToken cancellationToken) => CancellationToken = cancellationToken;
    }

    private sealed class RunPaintScope : IDisposable {
        private readonly RunPaintConversion? _previous;
        private bool _disposed;
        internal RunPaintScope(RunPaintConversion? previous) => _previous = previous;
        public void Dispose() {
            if (_disposed) return;
            RunPaintConversions.Value = _previous;
            _disposed = true;
        }
    }

    private static TableRunPaintState GetTableRunPaintState(Table table, Styles? styles) {
        RunPaintConversion? conversion = RunPaintConversions.Value;
        if (conversion == null) return new TableRunPaintState(table, styles, default);
        lock (conversion.Tables) {
            conversion.CancellationToken.ThrowIfCancellationRequested();
            if (!conversion.Tables.TryGetValue(table, out TableRunPaintState? state)) {
                state = new TableRunPaintState(table, styles, conversion.CancellationToken);
                conversion.Tables.Add(table, state);
            }
            return state;
        }
    }

    private sealed class TableRunPaintState {
        internal readonly Style[] Chain;
        internal readonly Style[] RunChain;
        internal readonly Dictionary<TableCell, (int Row, int Column, int Span)> Cells = new();
        internal readonly int Mask, HeaderRows, FooterStart, ColumnCount, RowCount;
        internal readonly int RowBandSize, ColumnBandSize;
        internal readonly bool HasConditionalFormatting;

        internal TableRunPaintState(Table table, Styles? styles, CancellationToken token) {
            string? styleId = table.GetFirstChild<TableProperties>()?.TableStyle?.Val?.Value;
            RunChain = EnumerateStyleChain(styles, styleId, StyleValues.Table).ToArray();
            // Paint retains its established built-in normal-table fallback boundary. Semantic
            // run formatting also honors explicitly authored properties in those definitions.
            Chain = RunChain.TakeWhile(style => !IsBuiltInNormalTableStyle(style.StyleId?.Value)).ToArray();
            RowBandSize = Chain.Select(style => style.StyleTableProperties?.GetFirstChild<TableStyleRowBandSize>()?.Val?.Value)
                .FirstOrDefault(size => size.HasValue) ?? 1;
            ColumnBandSize = Chain.Select(style => style.StyleTableProperties?.GetFirstChild<TableStyleColumnBandSize>()?.Val?.Value)
                .FirstOrDefault(size => size.HasValue) ?? 1;
            HasConditionalFormatting = RunChain.Any(style => style.Elements<TableStyleProperties>()
                .Any(conditional => conditional.GetFirstChild<RunPropertiesBaseStyle>()?.HasChildren == true ||
                    conditional.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>() is { } cellProperties &&
                    (cellProperties.GetFirstChild<Shading>() != null || cellProperties.GetFirstChild<TableCellBorders>() != null)));
            bool hasTableBorders = table.GetFirstChild<TableProperties>()?.TableBorders != null ||
                Chain.Any(style => style.StyleTableProperties?.TableBorders != null);
            if (!HasConditionalFormatting && !hasTableBorders) return;

            TableLook? look = table.GetFirstChild<TableProperties>()?.TableLook;
            Mask = WordTable.GetTableLookMask(look);
            TableRow[] rows = WordTableContent.Rows(table).ToArray();
            RowCount = rows.Length;
            bool inHeader = true;
            int maximumColumns = 0;
            for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
                token.ThrowIfCancellationRequested();
                TableRow row = rows[rowIndex];
                TableHeader? header = row.TableRowProperties?.GetFirstChild<TableHeader>();
                if (inHeader && header != null && WordTableRow.ReadOnOff(header)) HeaderRows++;
                else inHeader = false;
                int column = Math.Max(0, row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0);
                foreach (TableCell cell in WordTableContent.Cells(row)) {
                    if ((Cells.Count & 255) == 0) token.ThrowIfCancellationRequested();
                    Cells.Add(cell, (rowIndex, column, WordTableCellSpanResolver.GetColumnSpan(cell)));
                    column = checked(column + Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1));
                }
                column = checked(column + Math.Max(0, row.TableRowProperties?.GetFirstChild<GridAfter>()?.Val?.Value ?? 0));
                maximumColumns = Math.Max(maximumColumns, column);
            }
            if ((Mask & 0x0020) != 0) HeaderRows = Math.Max(HeaderRows, 1);
            FooterStart = (Mask & 0x0040) != 0 && rows.Length > HeaderRows ? rows.Length - 1 : rows.Length;
            int declaredColumns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0;
            ColumnCount = declaredColumns > 0 ? declaredColumns : maximumColumns;
        }
    }
}

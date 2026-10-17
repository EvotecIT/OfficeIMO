using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
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
        internal readonly Dictionary<TableCell, (int Row, int Column, int Span)> Cells = new();
        internal readonly int Mask, HeaderRows, FooterStart, ColumnCount;
        internal readonly bool HasConditionalPaint;

        internal TableRunPaintState(Table table, Styles? styles, CancellationToken token) {
            string? styleId = table.GetFirstChild<TableProperties>()?.TableStyle?.Val?.Value;
            Chain = EnumerateStyleChain(styles, styleId, StyleValues.Table)
                .TakeWhile(style => !IsBuiltInNormalTableStyle(style.StyleId?.Value)).ToArray();
            HasConditionalPaint = Chain.Any(style => style.Elements<TableStyleProperties>()
                .Any(conditional => HasRunBackground(conditional.GetFirstChild<RunPropertiesBaseStyle>())));
            if (!HasConditionalPaint) return;

            TableLook? look = table.GetFirstChild<TableProperties>()?.TableLook;
            Mask = look == null ? 0 : WordTable.GetTableLookMask(look);
            TableRow[] rows = table.Elements<TableRow>().ToArray();
            bool inHeader = true;
            int maximumColumns = 0;
            for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
                token.ThrowIfCancellationRequested();
                TableRow row = rows[rowIndex];
                TableHeader? header = row.TableRowProperties?.GetFirstChild<TableHeader>();
                if (inHeader && header != null && WordTableRow.ReadOnOff(header)) HeaderRows++;
                else inHeader = false;
                int column = Math.Max(0, row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0);
                foreach (TableCell cell in row.Elements<TableCell>()) {
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

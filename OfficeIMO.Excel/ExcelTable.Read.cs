using System.Diagnostics.CodeAnalysis;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Data;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelTable {
        /// <summary>
        /// Streams this table's data rows as typed models, using its column schema.
        /// Header and totals rows are excluded, including for tables with hidden headers.
        /// Enumerate while the owning workbook remains open; unsaved edits are included.
        /// </summary>
        /// <param name="options">Conversion, culture and safety options for reading the values.</param>
        /// <param name="cancellationToken">Cancels table discovery and row enumeration.</param>
        public IEnumerable<T> RowsAs<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicConstructors)] T>(
            ExcelReaderOptions? options = null,
            CancellationToken cancellationToken = default) {
            ExcelReaderOptions effective = options ?? new ExcelReaderOptions();
            using var linked = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken, effective.CancellationToken);
            CancellationToken token = linked.Token;
            token.ThrowIfCancellationRequested();
            var layout = Sheet.GetTableReadLayout(NameOrRange);
            if (layout.Range is null) yield break;

            using ExcelWorkbookDataReader reader = Sheet.CreateDataReader(
                effective.ForSheet(Sheet.Name, layout.Range, layout.HasHeaderRow, token));
            DataReaderMappingExtensions.GetConversionOptions(reader, out var culture, out var dateFormats,
                out var converter, out bool strict, out var errorPolicy);
            AutomaticRowMappingPlan<T> plan = AutomaticRowMappingPlan<T>.Create(layout.Columns, strict);
            while (reader.Read()) {
                token.ThrowIfCancellationRequested();
                T row = plan.MapReaderRow(reader, culture, dateFormats, converter, errorPolicy);
                token.ThrowIfCancellationRequested();
                yield return row;
            }
        }
    }

    public partial class ExcelSheet {
        // Resolve the native schema instead of using the first data row as headers.
        // A table wrapper can identify a table by its name, display name or range.
        internal (string? Range, bool HasHeaderRow, string[] Columns) GetTableReadLayout(string nameOrRange) =>
            Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                Table table = FindTableByRangeNameOrDisplayName(nameOrRange)
                    ?? throw new InvalidOperationException($"Table '{nameOrRange}' was not found on worksheet '{Name}'.");
                string? range = table.Reference?.Value;
                if (range is null || !ExcelA1.TryParseRange(range, out int firstRow, out int firstColumn,
                        out int lastRow, out int lastColumn)) {
                    throw new InvalidOperationException($"Table '{nameOrRange}' has no valid range.");
                }
                string[] columns = table.TableColumns?.Elements<TableColumn>()
                    .Select(column => column.Name?.Value ?? string.Empty).ToArray() ?? Array.Empty<string>();
                if (columns.Length != lastColumn - firstColumn + 1 || columns.Any(string.IsNullOrWhiteSpace)) {
                    throw new InvalidOperationException($"Table '{nameOrRange}' has no complete column schema.");
                }
                bool headers = (table.HeaderRowCount?.Value ?? 1U) > 0;
                uint totals = table.TotalsRowCount?.Value ?? 0U;
                long dataEnd = (long)lastRow - totals;
                string? readRange = dataEnd < firstRow ? null
                    : $"{ExcelA1.CellReference(firstRow, firstColumn)}:{ExcelA1.CellReference((int)dataEnd, lastColumn)}";
                return (readRange, headers, columns);
            });
    }
}

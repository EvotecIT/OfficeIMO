using System.Diagnostics.CodeAnalysis;
using System.Data;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Data;

namespace OfficeIMO.Excel {
    // Read convenience methods exposed directly on ExcelSheet to avoid separate reader usage.
    public partial class ExcelSheet {
        /// <summary>
        /// Returns the used range A1 address for this sheet.
        /// This is the canonical used-range API for an editable worksheet.
        /// </summary>
        public string UsedRangeA1 => GetUsedRangeA1();

        /// <summary>
        /// Creates a forward-only data reader over this worksheet in the current open workbook,
        /// including unsaved edits.
        /// </summary>
        public ExcelWorkbookDataReader CreateDataReader(ExcelReaderOptions? options = null) {
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            return _excelDocument.CreateDataReader(effectiveOptions.ForSheet(Name, effectiveOptions.A1Range, effectiveOptions.CancellationToken));
        }

        /// <summary>
        /// Streams the sheet's used range as instances of T using header-to-property mapping.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> remains open.
        /// </summary>
        /// <param name="options">Optional read options.</param>
        /// <param name="cancellationToken">Cancellation token observed during enumeration.</param>
        public IEnumerable<T> RowsAs<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicConstructors)] T>(
            ExcelReaderOptions? options = null,
            CancellationToken cancellationToken = default) {
            return RowsAsUsedRangeIterator<T>(options, cancellationToken);
        }

        /// <summary>
        /// Streams the sheet's used range as instances of T using explicit, AOT-friendly column assignments.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> remains open.
        /// </summary>
        /// <param name="configure">Configures the column assignments.</param>
        /// <param name="options">Optional read options.</param>
        /// <param name="cancellationToken">Cancellation token observed during enumeration.</param>
        public IEnumerable<T> RowsAs<T>(
            Action<RowMapper<T>> configure,
            ExcelReaderOptions? options = null,
            CancellationToken cancellationToken = default) where T : new() {
            if (configure is null) throw new ArgumentNullException(nameof(configure));
            return RowsAsUsedRangeIterator(configure, options, cancellationToken);
        }

        /// <summary>
        /// Streams the sheet's used range with a caller-supplied factory.
        /// This overload supports constructor-bound and other models without a public parameterless constructor.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> remains open.
        /// </summary>
        /// <param name="factory">Creates one model instance from the current row.</param>
        /// <param name="options">Optional read options.</param>
        /// <param name="cancellationToken">Cancellation token observed during enumeration.</param>
        public IEnumerable<T> RowsAs<T>(
            Func<IDataRecord, T> factory,
            ExcelReaderOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (factory is null) throw new ArgumentNullException(nameof(factory));
            return RowsAsUsedRangeIterator(factory, options, cancellationToken);
        }

        /// <summary>
        /// Streams the specified A1 range as instances of T using header-to-property mapping.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> remains open.
        /// </summary>
        public IEnumerable<T> RowsAs<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicConstructors)] T>(
            string a1Range,
            ExcelReaderOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            return RowsAsRangeIterator<T>(a1Range, options, cancellationToken);
        }

        /// <summary>
        /// Streams the specified A1 range as instances of T using explicit, AOT-friendly column assignments.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> remains open.
        /// </summary>
        /// <param name="a1Range">Inclusive A1 range containing the header and data rows.</param>
        /// <param name="configure">Configures the column assignments.</param>
        /// <param name="options">Optional read options.</param>
        /// <param name="cancellationToken">Cancellation token observed during enumeration.</param>
        public IEnumerable<T> RowsAs<T>(
            string a1Range,
            Action<RowMapper<T>> configure,
            ExcelReaderOptions? options = null,
            CancellationToken cancellationToken = default) where T : new() {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            if (configure is null) throw new ArgumentNullException(nameof(configure));
            return RowsAsRangeIterator(a1Range, configure, options, cancellationToken);
        }

        /// <summary>
        /// Streams the specified A1 range with a caller-supplied factory.
        /// This overload supports constructor-bound and other models without a public parameterless constructor.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> remains open.
        /// </summary>
        /// <param name="a1Range">Inclusive A1 range containing the header and data rows.</param>
        /// <param name="factory">Creates one model instance from the current row.</param>
        /// <param name="options">Optional read options.</param>
        /// <param name="cancellationToken">Cancellation token observed during enumeration.</param>
        public IEnumerable<T> RowsAs<T>(
            string a1Range,
            Func<IDataRecord, T> factory,
            ExcelReaderOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            if (factory is null) throw new ArgumentNullException(nameof(factory));
            return RowsAsRangeIterator(a1Range, factory, options, cancellationToken);
        }

        /// <summary>
        /// Streams the non-empty cells in this worksheet without probing every coordinate in its used range.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> is still open.
        /// </summary>
        /// <param name="options">Optional read options.</param>
        /// <param name="cancellationToken">Cancellation token observed during enumeration.</param>
        public IEnumerable<ExcelCellValueInfo> EnumerateCells(
            ExcelReaderOptions? options = null,
            CancellationToken cancellationToken = default) {
            return EnumerateCellsIterator(options, cancellationToken);
        }

        /// <summary>
        /// Streams the non-empty cells in the specified A1 range without probing every coordinate in the range.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> is still open.
        /// </summary>
        /// <param name="a1Range">Inclusive A1 range (for example, "A1:C100").</param>
        /// <param name="options">Optional read options.</param>
        /// <param name="cancellationToken">Cancellation token observed during enumeration.</param>
        public IEnumerable<ExcelCellValueInfo> EnumerateRange(
            string a1Range,
            ExcelReaderOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            return EnumerateRangeIterator(a1Range, options, cancellationToken);
        }

        /// <summary>
        /// Typed rows are data rows. When the read range ends on a table's totals row,
        /// that row is left out so a SUBTOTAL label or blank date does not become an object.
        /// </summary>
        private string? ExcludeTableTotalsRows(string? a1Range) {
            if (_worksheetPart.TableDefinitionParts.FirstOrDefault() == null) return a1Range;
            string range = a1Range ?? GetUsedRangeA1();
            if (!ExcelA1.TryParseRange(range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) return a1Range;
            int trimmedLastRow = lastRow;
            foreach (TableDefinitionPart part in _worksheetPart.TableDefinitionParts) {
                Table? table = part.Table;
                int totalsRows = (int)(table?.TotalsRowCount?.Value ?? 0U);
                if (table == null || totalsRows <= 0 ||
                    !ExcelA1.TryParseRange(table.Reference?.Value ?? string.Empty, out int tableFirstRow, out int tableFirstColumn, out int tableLastRow, out int tableLastColumn) ||
                    tableLastRow != trimmedLastRow || tableFirstColumn > lastColumn || tableLastColumn < firstColumn) {
                    continue;
                }

                int firstTotalsRow = tableLastRow - totalsRows + 1;
                if (firstTotalsRow > firstRow && firstTotalsRow > tableFirstRow) trimmedLastRow = firstTotalsRow - 1;
            }

            return trimmedLastRow == lastRow
                ? a1Range
                : ExcelA1.CellReference(firstRow, firstColumn) + ":" + ExcelA1.CellReference(trimmedLastRow, lastColumn);
        }

        private IEnumerable<T> RowsAsUsedRangeIterator<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicConstructors)] T>(
            ExcelReaderOptions? options,
            CancellationToken ct) {
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using ExcelWorkbookDataReader reader = _excelDocument.CreateDataReader(
                effectiveOptions.ForSheet(Name, a1Range: ExcludeTableTotalsRows(null), cancellationToken: token));
            foreach (T row in reader.RowsAs<T>()) {
                token.ThrowIfCancellationRequested();
                yield return row;
            }
        }

        private IEnumerable<T> RowsAsUsedRangeIterator<T>(
            Action<RowMapper<T>> configure,
            ExcelReaderOptions? options,
            CancellationToken cancellationToken) where T : new() {
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using ExcelWorkbookDataReader reader = _excelDocument.CreateDataReader(
                effectiveOptions.ForSheet(Name, a1Range: ExcludeTableTotalsRows(null), cancellationToken: token));
            foreach (T row in reader.RowsAs(configure)) {
                token.ThrowIfCancellationRequested();
                yield return row;
            }
        }

        private IEnumerable<T> RowsAsUsedRangeIterator<T>(
            Func<IDataRecord, T> factory,
            ExcelReaderOptions? options,
            CancellationToken cancellationToken) {
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using ExcelWorkbookDataReader reader = _excelDocument.CreateDataReader(
                effectiveOptions.ForSheet(Name, a1Range: ExcludeTableTotalsRows(null), cancellationToken: token));
            foreach (T row in reader.RowsAs(factory)) {
                token.ThrowIfCancellationRequested();
                yield return row;
            }
        }

        private IEnumerable<T> RowsAsRangeIterator<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicConstructors)] T>(
            string a1Range,
            ExcelReaderOptions? options,
            CancellationToken ct) {
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using ExcelWorkbookDataReader reader = _excelDocument.CreateDataReader(
                effectiveOptions.ForSheet(Name, ExcludeTableTotalsRows(a1Range), token));
            foreach (T row in reader.RowsAs<T>()) {
                token.ThrowIfCancellationRequested();
                yield return row;
            }
        }

        private IEnumerable<T> RowsAsRangeIterator<T>(
            string a1Range,
            Action<RowMapper<T>> configure,
            ExcelReaderOptions? options,
            CancellationToken cancellationToken) where T : new() {
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using ExcelWorkbookDataReader reader = _excelDocument.CreateDataReader(
                effectiveOptions.ForSheet(Name, ExcludeTableTotalsRows(a1Range), token));
            foreach (T row in reader.RowsAs(configure)) {
                token.ThrowIfCancellationRequested();
                yield return row;
            }
        }

        private IEnumerable<T> RowsAsRangeIterator<T>(
            string a1Range,
            Func<IDataRecord, T> factory,
            ExcelReaderOptions? options,
            CancellationToken cancellationToken) {
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using ExcelWorkbookDataReader reader = _excelDocument.CreateDataReader(
                effectiveOptions.ForSheet(Name, ExcludeTableTotalsRows(a1Range), token));
            foreach (T row in reader.RowsAs(factory)) {
                token.ThrowIfCancellationRequested();
                yield return row;
            }
        }

        private IEnumerable<ExcelCellValueInfo> EnumerateCellsIterator(
            ExcelReaderOptions? options,
            CancellationToken ct) {
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using var rdr = _excelDocument.CreateReader(effectiveOptions.WithCancellationToken(token));
            var sh = rdr.GetSheet(Name);
            foreach (ExcelCellValueInfo cell in sh.EnumerateCells(token)) {
                yield return cell;
            }
        }

        private IEnumerable<ExcelCellValueInfo> EnumerateRangeIterator(
            string a1Range,
            ExcelReaderOptions? options,
            CancellationToken ct) {
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using var rdr = _excelDocument.CreateReader(effectiveOptions.WithCancellationToken(token));
            var sh = rdr.GetSheet(Name);
            foreach (ExcelCellValueInfo cell in sh.EnumerateRange(a1Range, token)) {
                yield return cell;
            }
        }

    }
}

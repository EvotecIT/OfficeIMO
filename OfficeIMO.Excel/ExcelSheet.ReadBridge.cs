using System.Diagnostics.CodeAnalysis;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Read convenience methods exposed directly on ExcelSheet to avoid separate reader usage.
    /// </summary>
    public partial class ExcelSheet {
        /// <summary>
        /// Returns the used range A1 address for this sheet.
        /// Alias property for API ergonomics.
        /// </summary>
        public string UsedRangeA1 => GetUsedRangeA1();

        /// <summary>
        /// Creates a forward-only data reader over this worksheet in the current open workbook,
        /// including unsaved edits.
        /// </summary>
        public ExcelWorkbookDataReader CreateDataReader(ExcelReadOptions? options = null) {
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            return _excelDocument.CreateDataReader(effectiveOptions.ForSheet(Name, effectiveOptions.A1Range, effectiveOptions.CancellationToken));
        }

        /// <summary>
        /// Streams the sheet's used range as instances of T using header-to-property mapping.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> remains open.
        /// </summary>
        /// <param name="options">Optional read options.</param>
        /// <param name="ct">Cancellation token observed during enumeration.</param>
        public IEnumerable<T> RowsAs<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            ExcelReadOptions? options = null,
            CancellationToken ct = default) where T : new() {
            return RowsAsUsedRangeIterator<T>(options, ct);
        }

        /// <summary>
        /// Streams the specified A1 range as instances of T using header-to-property mapping.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> remains open.
        /// </summary>
        public IEnumerable<T> RowsAs<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            ExcelReadOptions? options = null,
            CancellationToken ct = default) where T : new() {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            return RowsAsRangeIterator<T>(a1Range, options, ct);
        }

        /// <summary>
        /// Streams the non-empty cells in this worksheet without probing every coordinate in its used range.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> is still open.
        /// </summary>
        /// <param name="options">Optional read options.</param>
        /// <param name="ct">Cancellation token observed during enumeration.</param>
        public IEnumerable<CellValueInfo> EnumerateCells(
            ExcelReadOptions? options = null,
            CancellationToken ct = default) {
            return EnumerateCellsIterator(options, ct);
        }

        /// <summary>
        /// Streams the non-empty cells in the specified A1 range without probing every coordinate in the range.
        /// Enumerate the returned sequence while the owning <see cref="ExcelDocument"/> is still open.
        /// </summary>
        /// <param name="a1Range">Inclusive A1 range (for example, "A1:C100").</param>
        /// <param name="options">Optional read options.</param>
        /// <param name="ct">Cancellation token observed during enumeration.</param>
        public IEnumerable<CellValueInfo> EnumerateRange(
            string a1Range,
            ExcelReadOptions? options = null,
            CancellationToken ct = default) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            return EnumerateRangeIterator(a1Range, options, ct);
        }

        private IEnumerable<T> RowsAsUsedRangeIterator<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            ExcelReadOptions? options,
            CancellationToken ct) where T : new() {
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using ExcelWorkbookDataReader reader = _excelDocument.CreateDataReader(
                effectiveOptions.ForSheet(Name, a1Range: null, cancellationToken: token));
            foreach (T row in reader.RowsAs<T>()) {
                token.ThrowIfCancellationRequested();
                yield return row;
            }
        }

        private IEnumerable<T> RowsAsRangeIterator<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            ExcelReadOptions? options,
            CancellationToken ct) where T : new() {
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using ExcelWorkbookDataReader reader = _excelDocument.CreateDataReader(
                effectiveOptions.ForSheet(Name, a1Range, token));
            foreach (T row in reader.RowsAs<T>()) {
                token.ThrowIfCancellationRequested();
                yield return row;
            }
        }

        private IEnumerable<CellValueInfo> EnumerateCellsIterator(
            ExcelReadOptions? options,
            CancellationToken ct) {
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using var rdr = _excelDocument.CreateReader(effectiveOptions.WithCancellationToken(token));
            var sh = rdr.GetSheet(Name);
            foreach (CellValueInfo cell in sh.EnumerateCells(token)) {
                yield return cell;
            }
        }

        private IEnumerable<CellValueInfo> EnumerateRangeIterator(
            string a1Range,
            ExcelReadOptions? options,
            CancellationToken ct) {
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using var rdr = _excelDocument.CreateReader(effectiveOptions.WithCancellationToken(token));
            var sh = rdr.GetSheet(Name);
            foreach (CellValueInfo cell in sh.EnumerateRange(a1Range, token)) {
                yield return cell;
            }
        }

    }
}

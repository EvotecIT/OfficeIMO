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
        /// Reads the sheet's used range as materialized instances of T using header-to-property mapping.
        /// </summary>
        /// <param name="options">Optional read options.</param>
        /// <param name="ct">Cancellation token observed during enumeration.</param>
        public IEnumerable<T> RowsAs<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            ExcelReadOptions? options = null,
            CancellationToken ct = default) where T : new() {
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using var rdr = _excelDocument.CreateReader(effectiveOptions.WithCancellationToken(token));
            var sh = rdr.GetSheet(this.Name);
            string a1Range = sh.GetUsedRangeA1(token);
            return sh.ReadObjects<T>(a1Range, ct: token).ToArray();
        }

        /// <summary>
        /// Reads the specified A1 range as materialized instances of T using header-to-property mapping.
        /// </summary>
        public IEnumerable<T> RowsAs<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            ExcelReadOptions? options = null,
            CancellationToken ct = default) where T : new() {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using var rdr = _excelDocument.CreateReader(effectiveOptions.WithCancellationToken(token));
            var sh = rdr.GetSheet(this.Name);
            return sh.ReadObjects<T>(a1Range, ct: token).ToArray();
        }

        /// <summary>
        /// Streams the sheet's used range as instances of T using header-to-property mapping.
        /// Enumerate the returned sequence while the owning ExcelDocument is still open.
        /// </summary>
        /// <param name="options">Optional read options.</param>
        /// <param name="ct">Cancellation token observed during enumeration.</param>
        public IEnumerable<T> RowsAsStream<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            ExcelReadOptions? options = null,
            CancellationToken ct = default) where T : new() {
            return RowsAsUsedRangeStreamIterator<T>(options, ct);
        }

        /// <summary>
        /// Streams the specified A1 range as instances of T using header-to-property mapping.
        /// Enumerate the returned sequence while the owning ExcelDocument is still open.
        /// </summary>
        /// <param name="a1Range">Inclusive A1 range (for example, "A1:C100").</param>
        /// <param name="options">Optional read options.</param>
        /// <param name="ct">Cancellation token observed during enumeration.</param>
        public IEnumerable<T> RowsAsStream<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            ExcelReadOptions? options = null,
            CancellationToken ct = default) where T : new() {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            return RowsAsRangeStreamIterator<T>(a1Range, options, ct);
        }

        private IEnumerable<T> RowsAsUsedRangeStreamIterator<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            ExcelReadOptions? options,
            CancellationToken ct) where T : new() {
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using var rdr = _excelDocument.CreateReader(effectiveOptions.WithCancellationToken(token));
            var sh = rdr.GetSheet(this.Name);
            string a1Range = sh.GetUsedRangeA1(token);
            foreach (var row in sh.ReadObjectsStream<T>(a1Range, token)) {
                yield return row;
            }
        }

        private IEnumerable<T> RowsAsRangeStreamIterator<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            ExcelReadOptions? options,
            CancellationToken ct) where T : new() {
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using var rdr = _excelDocument.CreateReader(effectiveOptions.WithCancellationToken(token));
            var sh = rdr.GetSheet(this.Name);
            foreach (var row in sh.ReadObjectsStream<T>(a1Range, token)) {
                yield return row;
            }
        }

    }
}

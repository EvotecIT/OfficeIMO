using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using OfficeIMO.Data;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Projects the used range in bounded parallel batches by matching headers to writable public properties.
        /// The worksheet reader remains single-owner and results retain source order.
        /// </summary>
        public IEnumerable<T> RowsAsParallel<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            ParallelRowMappingOptions? parallelOptions = null,
            ExcelReadOptions? options = null,
            CancellationToken cancellationToken = default) where T : new() =>
            RowsAsParallelUsedRangeIterator<T>(parallelOptions, options, cancellationToken);

        /// <summary>Projects the used range in bounded parallel batches with explicit, AOT-friendly assignments.</summary>
        public IEnumerable<T> RowsAsParallel<T>(
            Action<RowMapper<T>> configure,
            ParallelRowMappingOptions? parallelOptions = null,
            ExcelReadOptions? options = null,
            CancellationToken cancellationToken = default) where T : new() {
            if (configure is null) throw new ArgumentNullException(nameof(configure));
            return RowsAsParallelUsedRangeIterator(configure, parallelOptions, options, cancellationToken);
        }

        /// <summary>Projects the used range in bounded parallel batches with a caller-supplied factory.</summary>
        public IEnumerable<T> RowsAsParallel<T>(
            Func<IDataRecord, T> factory,
            ParallelRowMappingOptions? parallelOptions = null,
            ExcelReadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (factory is null) throw new ArgumentNullException(nameof(factory));
            return RowsAsParallelUsedRangeIterator(factory, parallelOptions, options, cancellationToken);
        }

        /// <summary>
        /// Projects an A1 range in bounded parallel batches by matching its first row to writable public properties.
        /// Results retain source order.
        /// </summary>
        public IEnumerable<T> RowsAsParallel<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            ParallelRowMappingOptions? parallelOptions = null,
            ExcelReadOptions? options = null,
            CancellationToken cancellationToken = default) where T : new() {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            return RowsAsParallelRangeIterator<T>(a1Range, parallelOptions, options, cancellationToken);
        }

        /// <summary>Projects an A1 range in bounded parallel batches with explicit, AOT-friendly assignments.</summary>
        public IEnumerable<T> RowsAsParallel<T>(
            string a1Range,
            Action<RowMapper<T>> configure,
            ParallelRowMappingOptions? parallelOptions = null,
            ExcelReadOptions? options = null,
            CancellationToken cancellationToken = default) where T : new() {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            if (configure is null) throw new ArgumentNullException(nameof(configure));
            return RowsAsParallelRangeIterator(a1Range, configure, parallelOptions, options, cancellationToken);
        }

        /// <summary>Projects an A1 range in bounded parallel batches with a caller-supplied factory.</summary>
        public IEnumerable<T> RowsAsParallel<T>(
            string a1Range,
            Func<IDataRecord, T> factory,
            ParallelRowMappingOptions? parallelOptions = null,
            ExcelReadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            if (factory is null) throw new ArgumentNullException(nameof(factory));
            return RowsAsParallelRangeIterator(a1Range, factory, parallelOptions, options, cancellationToken);
        }

        private IEnumerable<T> RowsAsParallelUsedRangeIterator<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            ParallelRowMappingOptions? parallelOptions,
            ExcelReadOptions? options,
            CancellationToken cancellationToken) where T : new() {
            using ExcelWorkbookDataReader reader = CreateParallelMappingReader(options, cancellationToken, out CancellationTokenSource linked);
            using (linked) {
                foreach (T row in reader.RowsAsParallel<T>(parallelOptions, linked.Token)) yield return row;
            }
        }

        private IEnumerable<T> RowsAsParallelUsedRangeIterator<T>(
            Action<RowMapper<T>> configure,
            ParallelRowMappingOptions? parallelOptions,
            ExcelReadOptions? options,
            CancellationToken cancellationToken) where T : new() {
            using ExcelWorkbookDataReader reader = CreateParallelMappingReader(options, cancellationToken, out CancellationTokenSource linked);
            using (linked) {
                foreach (T row in reader.RowsAsParallel(configure, parallelOptions, linked.Token)) yield return row;
            }
        }

        private IEnumerable<T> RowsAsParallelUsedRangeIterator<T>(
            Func<IDataRecord, T> factory,
            ParallelRowMappingOptions? parallelOptions,
            ExcelReadOptions? options,
            CancellationToken cancellationToken) {
            using ExcelWorkbookDataReader reader = CreateParallelMappingReader(options, cancellationToken, out CancellationTokenSource linked);
            using (linked) {
                foreach (T row in reader.RowsAsParallel(factory, parallelOptions, linked.Token)) yield return row;
            }
        }

        private IEnumerable<T> RowsAsParallelRangeIterator<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            ParallelRowMappingOptions? parallelOptions,
            ExcelReadOptions? options,
            CancellationToken cancellationToken) where T : new() {
            using ExcelWorkbookDataReader reader = CreateParallelMappingReader(options, cancellationToken, out CancellationTokenSource linked, a1Range);
            using (linked) {
                foreach (T row in reader.RowsAsParallel<T>(parallelOptions, linked.Token)) yield return row;
            }
        }

        private IEnumerable<T> RowsAsParallelRangeIterator<T>(
            string a1Range,
            Action<RowMapper<T>> configure,
            ParallelRowMappingOptions? parallelOptions,
            ExcelReadOptions? options,
            CancellationToken cancellationToken) where T : new() {
            using ExcelWorkbookDataReader reader = CreateParallelMappingReader(options, cancellationToken, out CancellationTokenSource linked, a1Range);
            using (linked) {
                foreach (T row in reader.RowsAsParallel(configure, parallelOptions, linked.Token)) yield return row;
            }
        }

        private IEnumerable<T> RowsAsParallelRangeIterator<T>(
            string a1Range,
            Func<IDataRecord, T> factory,
            ParallelRowMappingOptions? parallelOptions,
            ExcelReadOptions? options,
            CancellationToken cancellationToken) {
            using ExcelWorkbookDataReader reader = CreateParallelMappingReader(options, cancellationToken, out CancellationTokenSource linked, a1Range);
            using (linked) {
                foreach (T row in reader.RowsAsParallel(factory, parallelOptions, linked.Token)) yield return row;
            }
        }

        private ExcelWorkbookDataReader CreateParallelMappingReader(
            ExcelReadOptions? options,
            CancellationToken cancellationToken,
            out CancellationTokenSource linked,
            string? a1Range = null) {
            ExcelReadOptions effectiveOptions = options ?? new ExcelReadOptions();
            linked = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken,
                effectiveOptions.CancellationToken);
            try {
                return _excelDocument.CreateDataReader(
                    effectiveOptions.ForSheet(Name, a1Range, linked.Token));
            } catch {
                linked.Dispose();
                throw;
            }
        }
    }
}

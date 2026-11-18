namespace OfficeIMO.Excel {
    /// <summary>Applies a read's cancellation token to every advance of a result sequence.</summary>
    internal static class ExcelReadCancellation {
        /// <summary>
        /// Several read paths materialize their rows before returning them. Checking the token on each advance lets a
        /// caller cancel a small range, whose rows are already buffered, as promptly as a large streamed one.
        /// </summary>
        internal static IEnumerable<T> ObservePerItem<T>(IEnumerable<T> source, CancellationToken cancellationToken) =>
            cancellationToken.CanBeCanceled ? Observe(source, cancellationToken) : source;

        private static IEnumerable<T> Observe<T>(IEnumerable<T> source, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (T item in source) {
                cancellationToken.ThrowIfCancellationRequested();
                yield return item;
            }
        }
    }
}

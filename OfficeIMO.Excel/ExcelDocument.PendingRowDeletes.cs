using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private volatile PendingRowDeletes? _pendingRowDeletes;

        /// <summary>
        /// Row deletions accepted on a plain worksheet and not yet applied to its DOM.
        /// </summary>
        /// <remarks>
        /// Every route to the package goes through <see cref="_spreadSheetDocument"/>, the workbook part field or the
        /// part fields of <see cref="ExcelSheet"/>, which are properties so that no caller can skip the flush. Sheets
        /// holding objects that keep their own references into the package are never deferred.
        /// </remarks>
        internal sealed class PendingRowDeletes {
            internal PendingRowDeletes(ExcelSheet sheet) => Sheet = sheet;

            internal ExcelSheet Sheet { get; }
            internal List<(int FirstRow, int Count)> Operations { get; } = new List<(int FirstRow, int Count)>();
            internal List<int> DeletedRows { get; } = new List<int>();
            internal int OwnerThreadId { get; set; }

            /// <summary>Records a deletion given in current coordinates against the rows of the unchanged DOM.</summary>
            internal void Add(int firstRow, int count) {
                Operations.Add((firstRow, count));
                int original = firstRow;
                int index = 0;
                while (index < DeletedRows.Count && DeletedRows[index] <= original) {
                    original++;
                    index++;
                }
                for (int remaining = count; remaining > 0; remaining--) {
                    DeletedRows.Insert(index++, original++);
                    while (index < DeletedRows.Count && DeletedRows[index] == original) {
                        original++;
                        index++;
                    }
                }
            }
        }

        /// <summary>Marks the calling thread as the one extending the pending deletions, so its own checks do not flush them.</summary>
        internal readonly struct PendingRowDeleteScope : IDisposable {
            private readonly PendingRowDeletes? _pending;

            internal PendingRowDeleteScope(PendingRowDeletes? pending) => _pending = pending;

            public void Dispose() {
                if (_pending != null) _pending.OwnerThreadId = 0;
            }
        }

        internal PendingRowDeleteScope EnterPendingRowDeletes(ExcelSheet sheet) {
            PendingRowDeletes? pending = _pendingRowDeletes;
            if (pending == null) return default;
            if (!ReferenceEquals(pending.Sheet, sheet)) {
                FlushPendingRowDeletes();
                return default;
            }
            pending.OwnerThreadId = Environment.CurrentManagedThreadId;
            return new PendingRowDeleteScope(pending);
        }

        internal PendingRowDeletes? CurrentPendingRowDeletes => _pendingRowDeletes;

        internal PendingRowDeleteScope BeginPendingRowDeletes(ExcelSheet sheet, int firstRow, int count) {
            var pending = new PendingRowDeletes(sheet) { OwnerThreadId = Environment.CurrentManagedThreadId };
            pending.Add(firstRow, count);
            _pendingRowDeletes = pending;
            return new PendingRowDeleteScope(pending);
        }

        internal void FlushPendingRowDeletesIfAny() {
            if (_pendingRowDeletes != null) FlushPendingRowDeletes();
        }

        private void FlushPendingRowDeletes() {
            PendingRowDeletes? pending = _pendingRowDeletes;
            if (pending == null || pending.OwnerThreadId == Environment.CurrentManagedThreadId) return;
            lock (pending) {
                if (!ReferenceEquals(_pendingRowDeletes, pending)) return;
                pending.OwnerThreadId = Environment.CurrentManagedThreadId;
                try {
                    pending.Sheet.ApplyPendingRowDeletes(pending);
                } finally {
                    pending.OwnerThreadId = 0;
                    _pendingRowDeletes = null;
                }
            }
        }
    }
}

using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Shares the progress of the full XLSB validation pass with a reader that returns rows before it
    /// finishes. Validation publishes each completed row; the reader only decodes rows validation has
    /// finished, so every returned row passed the same checks the default reader runs up front.
    /// </summary>
    internal sealed class XlsbSinglePassGate {
        private readonly object _sync = new object();
        private readonly CancellationTokenSource _stop;
        private readonly CancellationToken _callerToken;
        private int _waiters;
        private int _stopped;
        private int _validatedLength;
        private int _completedRow = -1;
        private int _lastDataRow = -1;
        private int _firstOutOfSchemaRow = int.MaxValue;
        private bool _schemaFixed;
        private bool _complete;
        private Exception? _fault;
        private Task? _producer;
        private Task? _validation;
        private Func<Exception?>? _producerFault;

        internal XlsbSinglePassGate(CancellationToken cancellationToken) {
            _callerToken = cancellationToken;
            _stop = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        }

        /// <summary>Stops validation and inflation when the reader closes early.</summary>
        internal CancellationToken Token => _stop.Token;

        /// <summary>First column of the schema the header row fixed.</summary>
        internal int HeaderFirstColumn { get; private set; }

        /// <summary>Last column of the schema the header row fixed.</summary>
        internal int HeaderLastColumn { get; private set; }

        /// <summary>The header row, which is the first row with cells.</summary>
        internal int FirstDataRow { get; private set; }

        /// <summary>The complete validation result once validation has finished.</summary>
        internal XlsbWorksheetDiscovery Result { get; private set; }

        internal int ValidatedLength => Volatile.Read(ref _validatedLength);

        internal int LastDataRow => Volatile.Read(ref _lastDataRow);

        internal int FirstOutOfSchemaRow => Volatile.Read(ref _firstOutOfSchemaRow);

        internal bool IsComplete => Volatile.Read(ref _complete);

        internal void Attach(Task producer, Func<Exception?> producerFault, Task validation) {
            _producer = producer;
            _producerFault = producerFault;
            _validation = validation;
        }

        // Validation side ------------------------------------------------------------------------

        /// <summary>Fixes the schema once the header row is complete.</summary>
        internal void FixSchema(int firstColumn, int lastColumn, int headerRow) {
            HeaderFirstColumn = firstColumn;
            HeaderLastColumn = lastColumn;
            FirstDataRow = headerRow;
            Volatile.Write(ref _schemaFixed, true);
            Signal();
        }

        /// <summary>Records that a validated cell lies outside the fixed schema.</summary>
        internal void MarkOutOfSchema(int row) {
            if (row < _firstOutOfSchemaRow) Volatile.Write(ref _firstOutOfSchemaRow, row);
        }

        internal void ObserveDataRow(int row) => Volatile.Write(ref _lastDataRow, row);

        /// <summary>Publishes a validated prefix that ends after a row header or the end of sheet data.</summary>
        internal void PublishRowBoundary(int validatedLength, int completedRow) {
            Volatile.Write(ref _completedRow, completedRow);
            Volatile.Write(ref _validatedLength, validatedLength);
            Signal();
        }

        internal void Complete(XlsbWorksheetDiscovery result, int length) {
            Result = result;
            lock (_sync) {
                Volatile.Write(ref _validatedLength, length);
                Volatile.Write(ref _completedRow, int.MaxValue);
                Volatile.Write(ref _complete, true);
                Monitor.PulseAll(_sync);
            }
        }

        internal void Fail(Exception exception) {
            lock (_sync) {
                _fault = exception;
                Volatile.Write(ref _complete, true);
                Monitor.PulseAll(_sync);
            }
        }

        private void Signal() {
            // Pairs with the waiter's increment so a waiter either sees the new state or is pulsed.
            Interlocked.MemoryBarrier();
            if (Volatile.Read(ref _waiters) == 0) return;
            lock (_sync) Monitor.PulseAll(_sync);
        }

        // Reader side ----------------------------------------------------------------------------

        /// <summary>Waits until the header row fixed the schema or validation ended.</summary>
        internal bool WaitForSchema(CancellationToken cancellationToken) {
            Wait(static gate => Volatile.Read(ref gate._schemaFixed), cancellationToken);
            return Volatile.Read(ref _schemaFixed);
        }

        /// <summary>Waits until every cell of <paramref name="row"/> has been validated.</summary>
        internal void WaitForCompletedRow(int row, CancellationToken cancellationToken) {
            if (Volatile.Read(ref _completedRow) < row) WaitForCompletedRowSlow(row, cancellationToken);
        }

        /// <summary>Waits until validated bytes extend beyond <paramref name="position"/>; false at the end.</summary>
        internal bool WaitForBytesBeyond(int position, CancellationToken cancellationToken) =>
            Volatile.Read(ref _validatedLength) > position || WaitForBytesBeyondSlow(position, cancellationToken);

        /// <summary>Whether <paramref name="row"/> lies within the worksheet's data rows, waiting as needed.</summary>
        internal bool IncludesRow(int row, CancellationToken cancellationToken) =>
            Volatile.Read(ref _lastDataRow) >= row || IncludesRowSlow(row, cancellationToken);

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void WaitForCompletedRowSlow(int row, CancellationToken cancellationToken) {
            Wait(gate => Volatile.Read(ref gate._completedRow) >= row, cancellationToken);
            ThrowIfFaulted();
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool WaitForBytesBeyondSlow(int position, CancellationToken cancellationToken) {
            Wait(gate => Volatile.Read(ref gate._validatedLength) > position, cancellationToken);
            ThrowIfFaulted();
            return Volatile.Read(ref _validatedLength) > position;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool IncludesRowSlow(int row, CancellationToken cancellationToken) {
            Wait(gate => Volatile.Read(ref gate._lastDataRow) >= row, cancellationToken);
            // After a failure the row is treated as pending: the current row is still returned, and
            // the failure is raised by the next read, which waits for the pending row.
            return Volatile.Read(ref _lastDataRow) >= row || Volatile.Read(ref _complete) && _fault != null;
        }

        /// <summary>Waits for validation to end and raises its failure, if any.</summary>
        internal void WaitForCompletion(CancellationToken cancellationToken) {
            Wait(static _ => false, cancellationToken);
            // Inflation checks the part's exact length after validation has read every byte.
            try { _producer?.GetAwaiter().GetResult(); } catch { }
            // Cancellation reports the caller's token, as the default reader does.
            _callerToken.ThrowIfCancellationRequested();
            Exception? producerFault = _producerFault?.Invoke();
            if (producerFault != null) ExceptionDispatchInfo.Capture(producerFault).Throw();
            ThrowIfFaulted();
        }

        private void Wait(Func<XlsbSinglePassGate, bool> ready, CancellationToken cancellationToken) {
            if (ready(this) || Volatile.Read(ref _complete)) return;
            var spinner = new SpinWait();
            while (!spinner.NextSpinWillYield) {
                spinner.SpinOnce();
                if (ready(this) || Volatile.Read(ref _complete)) return;
            }
            lock (_sync) {
                Interlocked.Increment(ref _waiters);
                try {
                    while (!ready(this) && !_complete) {
                        cancellationToken.ThrowIfCancellationRequested();
                        Monitor.Wait(_sync, 50);
                    }
                } finally {
                    Interlocked.Decrement(ref _waiters);
                }
            }
            cancellationToken.ThrowIfCancellationRequested();
        }

        /// <summary>Raises the failure the default reader would raise: package errors before validation errors.</summary>
        internal void ThrowIfFaulted() {
            Exception? fault = Volatile.Read(ref _complete) ? _fault : null;
            if (fault == null) return;
            try { _producer?.GetAwaiter().GetResult(); } catch { }
            _callerToken.ThrowIfCancellationRequested();
            Exception? producerFault = _producerFault?.Invoke();
            ExceptionDispatchInfo.Capture(producerFault ?? fault).Throw();
        }

        /// <summary>Stops background work and waits for it, so the worksheet buffer can be returned.</summary>
        internal void StopAndJoin() {
            if (Interlocked.Exchange(ref _stopped, 1) != 0) return;
            _stop.Cancel();
            try { _validation?.GetAwaiter().GetResult(); } catch { }
            try { _producer?.GetAwaiter().GetResult(); } catch { }
            _stop.Dispose();
        }
    }
}

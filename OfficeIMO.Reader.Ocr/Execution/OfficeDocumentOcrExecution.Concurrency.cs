using System.Threading;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrExecutionExtensions {
    private sealed class TimedOutOcrOperationTracker {
        private int _hasTimedOutOperation;

        internal bool HasTimedOutOperation => Volatile.Read(ref _hasTimedOutOperation) != 0;

        internal void MarkTimedOut() {
            Interlocked.Exchange(ref _hasTimedOutOperation, 1);
        }
    }
}

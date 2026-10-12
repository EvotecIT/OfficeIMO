#nullable enable
#if NET8_0_OR_GREATER
using System.Runtime.ExceptionServices;
using System.Threading;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    private sealed class CsvAggregateFailure
    {
        private ExceptionDispatchInfo? _original;

        internal void Capture(Exception exception, CancellationToken stopToken)
        {
            // A sibling's owned stop check is a consequence, not the original failure.
            // Capture a callback/factory exception before the worker cancels that token.
            if (exception is OperationCanceledException canceled &&
                canceled.CancellationToken == stopToken && stopToken.IsCancellationRequested) return;
            Interlocked.CompareExchange(ref _original, ExceptionDispatchInfo.Capture(exception), null);
        }

        internal void RethrowOriginal() => Volatile.Read(ref _original)?.Throw();
    }
}
#endif

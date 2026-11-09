using System.Data.Common;
using System.Threading;

namespace OfficeIMO.Data;

public static partial class ParallelRowMappingExtensions {
    /// <summary>Keeps reader cancellation active while completed mapping batches are delivered.</summary>
    private static CancellationTokenSource? CreateMappingCancellationToken(
        DbDataReader reader, CancellationToken callerToken, out CancellationToken effectiveToken) {
        effectiveToken = callerToken;
        if (reader is not IDataReaderParallelBatchSource source) return null;
        CancellationToken readerToken = source.ParallelBatchCancellationToken;
        if (!readerToken.CanBeCanceled) return null;
        if (!callerToken.CanBeCanceled || callerToken == readerToken) {
            effectiveToken = readerToken;
            return null;
        }
        var linked = CancellationTokenSource.CreateLinkedTokenSource(readerToken, callerToken);
        effectiveToken = linked.Token;
        return linked;
    }
}

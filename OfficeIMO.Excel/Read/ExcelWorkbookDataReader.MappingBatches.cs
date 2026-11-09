using System.Data.Common;
using System.Threading;

namespace OfficeIMO.Excel;

public sealed partial class ExcelWorkbookDataReader : IDataReaderParallelBatchSource,
    IDataReaderTypedParallelBatchSource, IDataReaderParallelBatchInfo {
    private bool _isOwnedMappingBatch;
    private int _mappingBatchRowCount;

    bool IDataReaderParallelBatchSource.CanReadParallelBatches => false;
    bool IDataReaderTypedParallelBatchSource.CanReadTypedParallelBatches =>
        !_closed && !_isOwnedMappingBatch && _typeConverter is null &&
        CurrentReader is ExcelSchemaInferenceDataReader { CanCaptureMappingBatch: true };
    int IDataReaderParallelBatchSource.PreferredParallelBatchSize => 128;
    int IDataReaderParallelBatchInfo.ParallelBatchRowCount => _mappingBatchRowCount;

    bool IDataReaderParallelBatchSource.TryReadParallelBatch(int preferredBatchSize,
        CancellationToken cancellationToken, out DbDataReader? batchReader) {
        batchReader = null;
        _cancellationToken.ThrowIfCancellationRequested();
        cancellationToken.ThrowIfCancellationRequested();
        if (_closed || _isOwnedMappingBatch || _typeConverter is not null ||
            CurrentReader is not ExcelSchemaInferenceDataReader source ||
            !source.TryCaptureMappingBatch(preferredBatchSize, cancellationToken, out var batch, out int count))
            return false;
        if (batch is null) return true;
        try {
            batchReader = new ExcelWorkbookDataReader(new[] { _sheets[_resultIndex] },
                (_, _) => batch, batch, _culture, null, _requireAllColumnsMapped,
                _mappingErrorValuePolicy, CancellationToken.None) {
                _isOwnedMappingBatch = true, _mappingBatchRowCount = count
            };
            return true;
        } catch { batch.Dispose(); throw; }
    }
}

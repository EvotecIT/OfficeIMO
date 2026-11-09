using System.Data.Common;
using System.Threading;

namespace OfficeIMO.Excel;

public sealed partial class ExcelWorkbookDataReader : IDataReaderParallelBatchSource,
    IDataReaderTypedParallelBatchSource, IDataReaderParallelBatchInfo {
    private bool _isOwnedMappingBatch;
    private int _mappingBatchRowCount;

    CancellationToken IDataReaderParallelBatchSource.ParallelBatchCancellationToken => _cancellationToken;
    bool IDataReaderParallelBatchSource.CanReadParallelBatches => false;
    bool IDataReaderTypedParallelBatchSource.CanReadTypedParallelBatches =>
        !_closed && !_isOwnedMappingBatch && _typeConverter is null &&
        (CurrentReader is ExcelSchemaInferenceDataReader schema
            ? schema.CanCaptureMappingBatch
            : ExcelSheetReader.CanCaptureMappingBatch(CurrentReader));
    int IDataReaderParallelBatchSource.PreferredParallelBatchSize => 128;
    int IDataReaderParallelBatchInfo.ParallelBatchRowCount => _mappingBatchRowCount;

    bool IDataReaderParallelBatchSource.TryReadParallelBatch(int preferredBatchSize,
        CancellationToken cancellationToken, out DbDataReader? batchReader) {
        batchReader = null;
        _cancellationToken.ThrowIfCancellationRequested();
        cancellationToken.ThrowIfCancellationRequested();
        if (_closed || _isOwnedMappingBatch || _typeConverter is not null) return false;
        DbDataReader source = CurrentReader;
        DbDataReader? batch;
        int count;
        bool captured = source is ExcelSchemaInferenceDataReader schema
            ? schema.TryCaptureMappingBatch(preferredBatchSize, cancellationToken, out batch, out count)
            : ExcelSheetReader.TryCaptureMappingBatch(source, preferredBatchSize, cancellationToken, out batch, out count);
        if (!captured) return false;
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

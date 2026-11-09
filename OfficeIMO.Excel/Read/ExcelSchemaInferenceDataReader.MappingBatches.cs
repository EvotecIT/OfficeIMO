using System.Data.Common;
using System.Threading;

namespace OfficeIMO.Excel;

internal sealed partial class ExcelSchemaInferenceDataReader {
    private bool _mappingReplayOnly;

    internal bool CanCaptureMappingBatch => !_closed && !_mappingReplayOnly &&
        Array.IndexOf(_columnTypes, typeof(object)) < 0 &&
        ExcelSheetReader.CanCaptureMappingBatch(_inner);

    private ExcelSchemaInferenceDataReader(ExcelSchemaInferenceDataReader source,
        DbDataReader inner, List<object[]> samples, bool replayOnly) {
        _inner = inner;
        _culture = source._culture;
        _columnTypes = source._columnTypes;
        _sampledRows = samples;
        _mappingReplayOnly = replayOnly;
    }

    internal bool TryCaptureMappingBatch(int preferredBatchSize, CancellationToken ct,
        out DbDataReader? batch, out int count) {
        batch = null;
        count = 0;
        ct.ThrowIfCancellationRequested();
        if (!CanCaptureMappingBatch) return false;
        if (_sampleIndex < _sampledRows.Count) {
            count = Math.Min(Math.Max(1, preferredBatchSize), _sampledRows.Count - _sampleIndex);
            if (!ExcelSheetReader.TryCaptureMappingBatch(_inner, 0, ct, out var metadata, out _))
                return false;
            try {
                batch = new ExcelSchemaInferenceDataReader(this, metadata!,
                    _sampledRows.GetRange(_sampleIndex, count), replayOnly: true);
            } catch { metadata!.Dispose(); throw; }
            _sampleIndex += count;
            _sampledCurrentRow = null;
            SetUtf8TextCursor(false);
            return true;
        }
        if (!ExcelSheetReader.TryCaptureMappingBatch(_inner, preferredBatchSize, ct, out var rows, out count))
            return false;
        if (rows is not null) {
            try { batch = new ExcelSchemaInferenceDataReader(this, rows, new List<object[]>(), replayOnly: false); }
            catch { rows.Dispose(); throw; }
        }
        _sampledCurrentRow = null;
        SetUtf8TextCursor(false);
        return true;
    }
}

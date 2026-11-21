using System.Runtime.CompilerServices;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        // Present only while single-pass reading returns rows ahead of the validation pass.
        private readonly XlsbSinglePassGate? _gate;

        /// <summary>Frames the next record, waiting for single-pass validation to prove more of the part.</summary>
        private bool TryReadRecord(out XlsbRecordSlice record) {
            while (!_records.TryReadValidated(out record)) {
                if (!ExtendStreamedRecords(_records.Position)) return false;
            }
            return true;
        }

        /// <summary>Extends the readable prefix past <paramref name="position"/>; false once nothing follows.</summary>
        private bool ExtendStreamedRecords(int position) {
            if (_gate == null || !_gate.WaitForBytesBeyond(position, _activeReadCancellationToken)) return false;
            _records.ExtendTo(_gate.ValidatedLength);
            return true;
        }

        /// <summary>Whether a row past the data rows seen so far still holds data, waiting as needed.</summary>
        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool IncludesStreamedRow(int rowIndex) {
            if (_gate == null || !_gate.IncludesRow(rowIndex, _activeReadCancellationToken)) return false;
            _lastDataRow = _gate.LastDataRow;
            return true;
        }

        /// <summary>
        /// Waits until validation has proved every cell of the row, then reports a cell outside the
        /// header's columns: the default reader would have returned a wider table, so these rows cannot
        /// continue. A validation failure found first is raised instead, as the default reader raises it.
        /// </summary>
        private void PrepareStreamedRow(int rowIndex) {
            XlsbSinglePassGate gate = _gate!;
            gate.WaitForCompletedRow(rowIndex, _activeReadCancellationToken);
            if (rowIndex >= gate.FirstOutOfSchemaRow) {
                gate.WaitForCompletion(_activeReadCancellationToken);
                // The default reader rejects a table wider than the column limit before its first row.
                int firstColumn = Math.Min(_firstColumn, gate.Result.FirstColumn);
                int lastColumn = Math.Max(_firstColumn + FieldCount - 1, gate.Result.LastColumn);
                int fieldCount = checked(lastColumn - firstColumn + 1);
                if (fieldCount > _options.MaxDataReaderColumns) {
                    throw new InvalidDataException(
                        $"XLSB table column count {fieldCount} exceeds the configured limit of {_options.MaxDataReaderColumns}.");
                }
                throw new NotSupportedException(
                    $"XLSB row {rowIndex + 1} contains a cell outside the columns of the header row, which single-pass reading cannot return after earlier rows. Read the worksheet with EnableSinglePassReading disabled.");
            }
            int validated = gate.ValidatedLength;
            if (validated > _records.Length) _records.ExtendTo(validated);
        }
    }
}

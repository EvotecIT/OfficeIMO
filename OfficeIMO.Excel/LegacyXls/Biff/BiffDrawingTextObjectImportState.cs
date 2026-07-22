using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Pairs worksheet TxO drawing records with their following Continue-record text payload.
    /// </summary>
    internal sealed class BiffDrawingTextObjectImportState {
        private readonly string _sheetName;
        private readonly List<LegacyXlsDrawingRecord> _drawingRecords;
        private readonly LegacyXlsDecodedImageBudget? _decodedImageBudget;
        private PendingTextObject? _pendingTextObject;

        internal BiffDrawingTextObjectImportState(
            string sheetName,
            List<LegacyXlsDrawingRecord> drawingRecords,
            LegacyXlsDecodedImageBudget? decodedImageBudget = null) {
            _sheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
            _drawingRecords = drawingRecords ?? throw new ArgumentNullException(nameof(drawingRecords));
            _decodedImageBudget = decodedImageBudget;
        }

        internal bool TryReadTextObject(BiffRecord record) {
            if (_pendingTextObject != null) {
                return false;
            }

            if (!BiffDrawingMetadataReader.TryRead(record, _sheetName, out LegacyXlsDrawingRecord? drawingRecord, _decodedImageBudget)
                || drawingRecord?.TextObject == null) {
                return false;
            }

            _drawingRecords.Add(drawingRecord);
            LegacyXlsDrawingTextObject textObject = drawingRecord.TextObject;
            if (textObject.HasTextInContinueRecords || textObject.HasFormattingRunsInContinueRecords) {
                _pendingTextObject = new PendingTextObject(
                    _drawingRecords.Count - 1,
                    textObject,
                    new BiffTextObjectContinueReader(textObject.TextCharacterCount, textObject.FormattingRunByteCount));
            }

            return true;
        }

        internal bool TryReadContinue(byte[] payload) {
            if (_pendingTextObject == null) {
                return false;
            }

            if (!_pendingTextObject.Reader.TryRead(payload)) {
                _pendingTextObject = null;
                return true;
            }

            if (_pendingTextObject.Reader.Complete) {
                _drawingRecords[_pendingTextObject.DrawingRecordIndex] = _drawingRecords[_pendingTextObject.DrawingRecordIndex]
                    .WithTextObject(_pendingTextObject.TextObject.WithDecodedText(
                        _pendingTextObject.Reader.Text,
                        _pendingTextObject.Reader.FormattingRuns));
                _pendingTextObject = null;
            }

            return true;
        }

        internal void AddUnresolvedFeatures(
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            bool reportDiagnostics) {
            if (_pendingTextObject == null) {
                return;
            }

            LegacyXlsDrawingRecord drawingRecord = _drawingRecords[_pendingTextObject.DrawingRecordIndex];
            LegacyXlsUnsupportedFeature feature = BiffUnsupportedRecordDiagnostics.CreateUnsupportedRecordFeature(
                drawingRecord.RecordType,
                drawingRecord.RecordOffset,
                _sheetName);
            unsupportedFeatures.Add(feature);
            if (BiffUnsupportedRecordDiagnostics.TryCreatePreservedFeatureRecord(feature, drawingRecord.PayloadLength, out LegacyXlsPreservedFeatureRecord? preservedRecord)) {
                preservedFeatureRecords.Add(preservedRecord!);
            }

            if (reportDiagnostics) {
                BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(
                    diagnostics,
                    drawingRecord.RecordType,
                    drawingRecord.RecordOffset,
                    _sheetName);
            }

            _pendingTextObject = null;
        }

        private sealed class PendingTextObject {
            internal PendingTextObject(
                int drawingRecordIndex,
                LegacyXlsDrawingTextObject textObject,
                BiffTextObjectContinueReader reader) {
                DrawingRecordIndex = drawingRecordIndex;
                TextObject = textObject;
                Reader = reader;
            }

            internal int DrawingRecordIndex { get; }

            internal LegacyXlsDrawingTextObject TextObject { get; }

            internal BiffTextObjectContinueReader Reader { get; }
        }
    }
}

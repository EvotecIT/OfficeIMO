using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Reassembles MsoDrawing OfficeArt payloads that continue into following BIFF Continue records.
    /// </summary>
    internal sealed class BiffDrawingContinuationImportState {
        private readonly string _sheetName;
        private readonly List<LegacyXlsDrawingRecord> _drawingRecords;
        private readonly LegacyXlsDecodedImageBudget? _decodedImageBudget;
        private PendingDrawing? _pendingDrawing;

        internal BiffDrawingContinuationImportState(
            string sheetName,
            List<LegacyXlsDrawingRecord> drawingRecords,
            LegacyXlsDecodedImageBudget? decodedImageBudget = null) {
            _sheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
            _drawingRecords = drawingRecords ?? throw new ArgumentNullException(nameof(drawingRecords));
            _decodedImageBudget = decodedImageBudget;
        }

        internal bool TryReadDrawing(
            BiffRecord record,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            bool reportDiagnostics) {
            if (_pendingDrawing != null || record.Type != (ushort)BiffRecordType.Drawing) {
                return false;
            }

            if (!BiffDrawingMetadataReader.TryRead(record, _sheetName, out LegacyXlsDrawingRecord? drawingRecord, _decodedImageBudget)) {
                return false;
            }

            _drawingRecords.Add(drawingRecord!);
            if (!TryGetRequiredPayloadLength(record.Payload, out int requiredPayloadLength)
                || requiredPayloadLength <= record.Payload.Length) {
                if (!drawingRecord!.HasSupportedDrawingMetadata) {
                    AddUnsupportedFeature(
                        unsupportedFeatures,
                        preservedFeatureRecords,
                        diagnostics,
                        reportDiagnostics,
                        record.Type,
                        record.Offset,
                        record.Payload.Length);
                }

                return true;
            }

            _pendingDrawing = new PendingDrawing(
                _drawingRecords.Count - 1,
                record.Type,
                record.Offset,
                requiredPayloadLength,
                record.Payload);
            return true;
        }

        internal bool TryReadContinue(
            byte[] payload,
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            bool reportDiagnostics,
            out BiffRecord? assembledRecord) {
            assembledRecord = null;
            if (_pendingDrawing == null) {
                return false;
            }

            _pendingDrawing.Append(payload);
            if (!_pendingDrawing.Complete) {
                return true;
            }

            byte[] assembledPayload = _pendingDrawing.GetAssembledPayload();
            var completedRecord = new BiffRecord(_pendingDrawing.RecordType, _pendingDrawing.RecordOffset, assembledPayload);
            if (BiffDrawingMetadataReader.TryRead(completedRecord, _sheetName, out LegacyXlsDrawingRecord? drawingRecord, _decodedImageBudget)) {
                _drawingRecords[_pendingDrawing.DrawingRecordIndex] = drawingRecord!;
                if (!drawingRecord!.HasSupportedDrawingMetadata) {
                    AddUnsupportedFeature(
                        unsupportedFeatures,
                        preservedFeatureRecords,
                        diagnostics,
                        reportDiagnostics,
                        _pendingDrawing.RecordType,
                        _pendingDrawing.RecordOffset,
                        assembledPayload.Length);
                }
            } else {
                AddUnsupportedFeature(
                    unsupportedFeatures,
                    preservedFeatureRecords,
                    diagnostics,
                    reportDiagnostics,
                    _pendingDrawing.RecordType,
                    _pendingDrawing.RecordOffset,
                    assembledPayload.Length);
            }

            assembledRecord = completedRecord;
            _pendingDrawing = null;
            return true;
        }

        internal void AddUnresolvedFeatures(
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            bool reportDiagnostics) {
            if (_pendingDrawing == null) {
                return;
            }

            LegacyXlsDrawingRecord drawingRecord = _drawingRecords[_pendingDrawing.DrawingRecordIndex];
            if (!drawingRecord.HasSupportedDrawingMetadata) {
                AddUnsupportedFeature(
                    unsupportedFeatures,
                    preservedFeatureRecords,
                    diagnostics,
                    reportDiagnostics,
                    _pendingDrawing.RecordType,
                    _pendingDrawing.RecordOffset,
                    _pendingDrawing.AvailableLength);
            }

            _pendingDrawing = null;
        }

        internal static bool RequiresContinuation(BiffRecord record) {
            return record.Type == (ushort)BiffRecordType.Drawing
                && TryGetRequiredPayloadLength(record.Payload, out int requiredPayloadLength)
                && requiredPayloadLength > record.Payload.Length;
        }

        private static bool TryGetRequiredPayloadLength(byte[] payload, out int requiredPayloadLength) {
            requiredPayloadLength = 0;
            if (payload.Length < 8) {
                return false;
            }

            uint escherPayloadLength = BiffRecordReader.ReadUInt32(payload, 4);
            if (escherPayloadLength > int.MaxValue - 8) {
                return false;
            }

            requiredPayloadLength = checked(8 + (int)escherPayloadLength);
            return true;
        }

        private void AddUnsupportedFeature(
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            bool reportDiagnostics,
            ushort recordType,
            int recordOffset,
            int payloadLength) {
            LegacyXlsUnsupportedFeature feature = BiffUnsupportedRecordDiagnostics.CreateUnsupportedRecordFeature(
                recordType,
                recordOffset,
                _sheetName);
            unsupportedFeatures.Add(feature);
            if (BiffUnsupportedRecordDiagnostics.TryCreatePreservedFeatureRecord(feature, payloadLength, out LegacyXlsPreservedFeatureRecord? preservedRecord)) {
                preservedFeatureRecords.Add(preservedRecord!);
            }

            if (reportDiagnostics) {
                BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(
                    diagnostics,
                    recordType,
                    recordOffset,
                    _sheetName);
            }
        }

        private sealed class PendingDrawing {
            private readonly MemoryStream _payload;

            internal PendingDrawing(
                int drawingRecordIndex,
                ushort recordType,
                int recordOffset,
                int requiredLength,
                byte[] initialPayload) {
                DrawingRecordIndex = drawingRecordIndex;
                RecordType = recordType;
                RecordOffset = recordOffset;
                RequiredLength = requiredLength;
                _payload = new MemoryStream(requiredLength);
                Append(initialPayload);
            }

            internal int DrawingRecordIndex { get; }

            internal ushort RecordType { get; }

            internal int RecordOffset { get; }

            internal int RequiredLength { get; }

            internal int AvailableLength => checked((int)_payload.Length);

            internal bool Complete => _payload.Length >= RequiredLength;

            internal void Append(byte[] payload) {
                if (payload == null) {
                    throw new ArgumentNullException(nameof(payload));
                }

                int bytesNeeded = RequiredLength - checked((int)_payload.Length);
                if (bytesNeeded <= 0) {
                    return;
                }

                int bytesToWrite = Math.Min(bytesNeeded, payload.Length);
                _payload.Write(payload, 0, bytesToWrite);
            }

            internal byte[] GetAssembledPayload() {
                byte[] bytes = _payload.ToArray();
                if (bytes.Length == RequiredLength) {
                    return bytes;
                }

                byte[] trimmed = new byte[RequiredLength];
                Buffer.BlockCopy(bytes, 0, trimmed, 0, Math.Min(bytes.Length, trimmed.Length));
                return trimmed;
            }
        }
    }
}

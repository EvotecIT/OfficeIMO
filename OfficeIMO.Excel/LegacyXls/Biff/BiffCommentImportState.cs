using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Pairs BIFF NOTE metadata with note-type OBJ/TXO text records for worksheet comments.
    /// </summary>
    internal sealed class BiffCommentImportState {
        private readonly LegacyXlsWorksheet _sheet;
        private readonly LegacyXlsDecodedImageBudget? _decodedImageBudget;
        private readonly Dictionary<ushort, PendingNote> _notes = new();
        private readonly Dictionary<ushort, string> _texts = new();
        private readonly Dictionary<ushort, IReadOnlyList<LegacyXlsCommentFormattingRun>> _formattingRuns = new();
        private readonly Dictionary<ushort, CommentObjectInfo> _objectInfos = new();
        private readonly Queue<LegacyXlsDrawingAnchor> _pendingAnchors = new();
        private readonly Queue<PendingDrawingRecord> _pendingDrawingRecords = new();
        private readonly HashSet<int> _drawingAnchorRecordOffsets = new();
        private readonly HashSet<ushort> _imported = new();
        private PendingTextObject? _pendingTextObject;
        private ushort? _pendingCommentObjectId;

        internal BiffCommentImportState(
            LegacyXlsWorksheet sheet,
            LegacyXlsDecodedImageBudget? decodedImageBudget = null) {
            _sheet = sheet;
            _decodedImageBudget = decodedImageBudget;
        }

        internal bool TryReadObject(byte[] payload) {
            _pendingCommentObjectId = null;
            if (payload.Length < 8) {
                return false;
            }

            ushort ft = BiffRecordReader.ReadUInt16(payload, 0);
            ushort cb = BiffRecordReader.ReadUInt16(payload, 2);
            ushort objectType = BiffRecordReader.ReadUInt16(payload, 4);
            if (ft != 0x0015 || cb != 0x0012 || objectType != 0x0019) {
                return false;
            }

            _pendingCommentObjectId = BiffRecordReader.ReadUInt16(payload, 6);
            ushort? objectFlags = null;
            if (cb >= 6 && payload.Length >= 10) {
                objectFlags = BiffRecordReader.ReadUInt16(payload, 8);
            }

            LegacyXlsDrawingAnchor? anchor = _pendingAnchors.Count > 0 ? _pendingAnchors.Dequeue() : null;
            if (_pendingDrawingRecords.Count > 0) {
                _pendingDrawingRecords.Dequeue();
            }

            _objectInfos[_pendingCommentObjectId.Value] = new CommentObjectInfo(objectType, objectFlags, anchor);
            return true;
        }

        internal bool TryReadDrawingAnchors(BiffRecord record, out LegacyXlsDrawingRecord? drawingRecord) {
            drawingRecord = null;
            if (BiffDrawingMetadataReader.TryRead(record, _sheet.Name, out LegacyXlsDrawingRecord? parsedRecord, _decodedImageBudget) && parsedRecord!.AnchorEntries.Count > 0) {
                if (!_drawingAnchorRecordOffsets.Add(record.Offset)) {
                    drawingRecord = parsedRecord;
                    return false;
                }

                drawingRecord = parsedRecord;
                foreach (LegacyXlsDrawingAnchor anchor in parsedRecord.AnchorEntries) {
                    _pendingAnchors.Enqueue(anchor);
                }

                _pendingDrawingRecords.Enqueue(new PendingDrawingRecord(
                    record.Type,
                    record.Offset,
                    record.Payload.Length,
                    parsedRecord.HasSupportedDrawingMetadata));
                return true;
            }

            return false;
        }

        internal void AddPendingDrawingFeatures(
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            bool reportDiagnostics) {
            while (_pendingDrawingRecords.Count > 0) {
                PendingDrawingRecord pendingDrawing = _pendingDrawingRecords.Dequeue();
                if (pendingDrawing.HasSupportedOfficeArtMetadata) {
                    continue;
                }

                LegacyXlsUnsupportedFeature feature = BiffUnsupportedRecordDiagnostics.CreateUnsupportedRecordFeature(
                    pendingDrawing.RecordType,
                    pendingDrawing.RecordOffset,
                    _sheet.Name);
                unsupportedFeatures.Add(feature);
                if (BiffUnsupportedRecordDiagnostics.TryCreatePreservedFeatureRecord(feature, pendingDrawing.RecordPayloadLength, out LegacyXlsPreservedFeatureRecord? preservedRecord)) {
                    preservedFeatureRecords.Add(preservedRecord!);
                }

                if (reportDiagnostics) {
                    BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(
                        diagnostics,
                        pendingDrawing.RecordType,
                        pendingDrawing.RecordOffset,
                        _sheet.Name);
                }
            }

            _pendingAnchors.Clear();
        }

        internal bool TryReadTextObject(byte[] payload) {
            if (!_pendingCommentObjectId.HasValue) {
                return false;
            }

            _pendingTextObject = null;
            if (payload.Length < 16) {
                return true;
            }

            ushort textCharacters = BiffRecordReader.ReadUInt16(payload, 10);
            ushort formattingRunBytes = BiffRecordReader.ReadUInt16(payload, 12);
            if (textCharacters == 0) {
                return true;
            }

            _pendingTextObject = new PendingTextObject(_pendingCommentObjectId.Value, textCharacters, formattingRunBytes);
            _pendingCommentObjectId = null;
            return true;
        }

        internal bool TryReadContinue(byte[] payload) {
            if (_pendingTextObject == null) {
                return false;
            }

            if (!_pendingTextObject.TryRead(payload)) {
                _pendingTextObject = null;
                return true;
            }

            if (_pendingTextObject.Complete) {
                _texts[_pendingTextObject.ObjectId] = _pendingTextObject.Text;
                _formattingRuns[_pendingTextObject.ObjectId] = _pendingTextObject.FormattingRuns;
                TryAddComment(_pendingTextObject.ObjectId);
                _pendingTextObject = null;
            }

            return true;
        }

        internal bool TryReadNote(byte[] payload, int recordOffset) {
            if (payload.Length < 10) {
                return false;
            }

            ushort row = BiffRecordReader.ReadUInt16(payload, 0);
            ushort column = BiffRecordReader.ReadUInt16(payload, 2);
            ushort flags = BiffRecordReader.ReadUInt16(payload, 4);
            ushort objectId = BiffRecordReader.ReadUInt16(payload, 6);
            int offset = 8;
            string author = BiffStringReader.ReadUnicodeString(payload, ref offset);
            if (string.IsNullOrWhiteSpace(author)) {
                author = "OfficeIMO";
            }

            _notes[objectId] = new PendingNote(row + 1, column + 1, author, (flags & 0x0002) != 0, recordOffset, payload.Length);
            TryAddComment(objectId);
            return true;
        }

        internal void AddUnresolvedFeatures(
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            bool reportDiagnostics) {
            foreach (KeyValuePair<ushort, PendingNote> entry in _notes) {
                if (!_imported.Contains(entry.Key)) {
                    LegacyXlsUnsupportedFeature feature = BiffUnsupportedRecordDiagnostics.CreateUnsupportedRecordFeature(
                        (ushort)BiffRecordType.Note,
                        entry.Value.RecordOffset,
                        _sheet.Name);
                    unsupportedFeatures.Add(feature);
                    if (BiffUnsupportedRecordDiagnostics.TryCreatePreservedFeatureRecord(feature, entry.Value.RecordPayloadLength, out LegacyXlsPreservedFeatureRecord? preservedRecord)) {
                        preservedFeatureRecords.Add(preservedRecord!);
                    }

                    if (reportDiagnostics) {
                        BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(
                            diagnostics,
                            (ushort)BiffRecordType.Note,
                            entry.Value.RecordOffset,
                            _sheet.Name);
                    }
                }
            }
        }

        private void TryAddComment(ushort objectId) {
            if (_imported.Contains(objectId)
                || !_notes.TryGetValue(objectId, out PendingNote note)
                || !_texts.TryGetValue(objectId, out string? text)
                || string.IsNullOrEmpty(text)) {
                return;
            }

            _formattingRuns.TryGetValue(objectId, out IReadOnlyList<LegacyXlsCommentFormattingRun>? runs);
            _objectInfos.TryGetValue(objectId, out CommentObjectInfo objectInfo);
            _sheet.AddComment(new LegacyXlsComment(
                note.Row,
                note.Column,
                text!,
                note.Author,
                objectId,
                note.Visible,
                runs,
                objectInfo.ObjectType,
                objectInfo.ObjectFlags,
                objectInfo.Anchor));
            _imported.Add(objectId);
        }

        private readonly struct CommentObjectInfo {
            internal CommentObjectInfo(ushort? objectType, ushort? objectFlags, LegacyXlsDrawingAnchor? anchor) {
                ObjectType = objectType;
                ObjectFlags = objectFlags;
                Anchor = anchor;
            }

            internal ushort? ObjectType { get; }

            internal ushort? ObjectFlags { get; }

            internal LegacyXlsDrawingAnchor? Anchor { get; }
        }

        private readonly struct PendingNote {
            internal PendingNote(int row, int column, string author, bool visible, int recordOffset, int recordPayloadLength) {
                Row = row;
                Column = column;
                Author = author;
                Visible = visible;
                RecordOffset = recordOffset;
                RecordPayloadLength = recordPayloadLength;
            }

            internal int Row { get; }

            internal int Column { get; }

            internal string Author { get; }

            internal bool Visible { get; }

            internal int RecordOffset { get; }

            internal int RecordPayloadLength { get; }
    }

        private readonly struct PendingDrawingRecord {
            internal PendingDrawingRecord(ushort recordType, int recordOffset, int recordPayloadLength, bool hasSupportedOfficeArtMetadata) {
                RecordType = recordType;
                RecordOffset = recordOffset;
                RecordPayloadLength = recordPayloadLength;
                HasSupportedOfficeArtMetadata = hasSupportedOfficeArtMetadata;
            }

            internal ushort RecordType { get; }

            internal int RecordOffset { get; }

            internal int RecordPayloadLength { get; }

            internal bool HasSupportedOfficeArtMetadata { get; }
        }

        private sealed class PendingTextObject {
            private readonly BiffTextObjectContinueReader _reader;

            internal PendingTextObject(ushort objectId, int textCharacters, int formattingRunBytes) {
                ObjectId = objectId;
                _reader = new BiffTextObjectContinueReader(textCharacters, formattingRunBytes);
            }

            internal ushort ObjectId { get; }

            internal string Text => _reader.Text;

            internal IReadOnlyList<LegacyXlsCommentFormattingRun> FormattingRuns => _reader.FormattingRuns;

            internal bool Complete => _reader.Complete;

            internal bool TryRead(byte[] payload) {
                return _reader.TryRead(payload);
            }
        }
    }
}

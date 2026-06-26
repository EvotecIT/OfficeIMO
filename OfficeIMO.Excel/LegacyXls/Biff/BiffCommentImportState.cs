using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Pairs BIFF NOTE metadata with note-type OBJ/TXO text records for worksheet comments.
    /// </summary>
    internal sealed class BiffCommentImportState {
        private readonly LegacyXlsWorksheet _sheet;
        private readonly Dictionary<ushort, PendingNote> _notes = new();
        private readonly Dictionary<ushort, string> _texts = new();
        private readonly Dictionary<ushort, IReadOnlyList<LegacyXlsCommentFormattingRun>> _formattingRuns = new();
        private readonly Dictionary<ushort, CommentObjectInfo> _objectInfos = new();
        private readonly Queue<LegacyXlsDrawingAnchor> _pendingAnchors = new();
        private readonly HashSet<ushort> _imported = new();
        private PendingTextObject? _pendingTextObject;
        private ushort? _pendingCommentObjectId;

        internal BiffCommentImportState(LegacyXlsWorksheet sheet) {
            _sheet = sheet;
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
            _objectInfos[_pendingCommentObjectId.Value] = new CommentObjectInfo(objectType, objectFlags, anchor);
            return true;
        }

        internal void TryReadDrawingAnchors(BiffRecord record) {
            if (BiffDrawingMetadataReader.TryReadClientAnchors(record, out IReadOnlyList<LegacyXlsDrawingAnchor> anchors)) {
                foreach (LegacyXlsDrawingAnchor anchor in anchors) {
                    _pendingAnchors.Enqueue(anchor);
                }
            }
        }

        internal void DiscardPendingDrawingAnchors() {
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

        private sealed class PendingTextObject {
            private readonly StringBuilder _builder = new();
            private readonly List<byte> _formattingBytes = new();
            private int _remainingCharacters;
            private int _remainingFormattingBytes;
            private IReadOnlyList<LegacyXlsCommentFormattingRun>? _formattingRuns;

            internal PendingTextObject(ushort objectId, int textCharacters, int formattingRunBytes) {
                ObjectId = objectId;
                _remainingCharacters = textCharacters;
                _remainingFormattingBytes = formattingRunBytes;
            }

            internal ushort ObjectId { get; }

            internal string Text => _builder.ToString();

            internal IReadOnlyList<LegacyXlsCommentFormattingRun> FormattingRuns => _formattingRuns ??= ParseFormattingRuns();

            internal bool Complete => _remainingCharacters == 0 && _remainingFormattingBytes == 0;

            internal bool TryRead(byte[] payload) {
                int offset = 0;
                if (_remainingCharacters > 0) {
                    if (payload.Length == 0) {
                        return false;
                    }

                    byte options = payload[offset++];
                    bool isUtf16 = (options & 0x01) != 0;
                    int bytesPerCharacter = isUtf16 ? 2 : 1;
                    int availableCharacters = (payload.Length - offset) / bytesPerCharacter;
                    int charactersToRead = Math.Min(_remainingCharacters, availableCharacters);
                    int bytesToRead = checked(charactersToRead * bytesPerCharacter);
                    if (isUtf16) {
                        _builder.Append(Encoding.Unicode.GetString(payload, offset, bytesToRead));
                    } else {
                        for (int i = 0; i < bytesToRead; i++) {
                            _builder.Append((char)payload[offset + i]);
                        }
                    }

                    offset += bytesToRead;
                    _remainingCharacters -= charactersToRead;
                }

                if (_remainingCharacters == 0 && _remainingFormattingBytes > 0) {
                    int formattingBytesToRead = Math.Min(_remainingFormattingBytes, payload.Length - offset);
                    for (int i = 0; i < formattingBytesToRead; i++) {
                        _formattingBytes.Add(payload[offset + i]);
                    }

                    _remainingFormattingBytes -= formattingBytesToRead;
                }

                return true;
            }

            private IReadOnlyList<LegacyXlsCommentFormattingRun> ParseFormattingRuns() {
                if (_formattingBytes.Count < 16 || _formattingBytes.Count % 8 != 0) {
                    return Array.Empty<LegacyXlsCommentFormattingRun>();
                }

                int runCount = (_formattingBytes.Count / 8) - 1;
                if (runCount <= 0) {
                    return Array.Empty<LegacyXlsCommentFormattingRun>();
                }

                var runs = new List<LegacyXlsCommentFormattingRun>(runCount);
                byte[] bytes = _formattingBytes.ToArray();
                for (int i = 0; i < runCount; i++) {
                    int offset = i * 8;
                    ushort startCharacter = BiffRecordReader.ReadUInt16(bytes, offset);
                    ushort fontIndex = BiffRecordReader.ReadUInt16(bytes, offset + 2);
                    if (startCharacter > Text.Length) {
                        continue;
                    }

                    if (runs.Count > 0 && startCharacter <= runs[runs.Count - 1].StartCharacter) {
                        continue;
                    }

                    runs.Add(new LegacyXlsCommentFormattingRun(startCharacter, fontIndex));
                }

                return runs;
            }
        }
    }
}

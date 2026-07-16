using OfficeIMO.Excel.LegacyXls.Model;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private static byte[] BuildLabelSstPayload(ushort row, ushort column, ushort styleIndex, uint sharedStringIndex) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, row);
            WriteUInt16(stream, column);
            WriteUInt16(stream, styleIndex);
            WriteUInt32(stream, sharedStringIndex);
            return stream.ToArray();
        }

        private static void WriteFormulaCachedStringRecords(Stream stream, string text) {
            var writer = new LegacyXlsStringSegmentWriter();
            _ = writer.WriteUnicodeString(text, Array.Empty<LegacyXlsTextFormattingRun>());
            IReadOnlyList<byte[]> payloads = writer.GetPayloads();
            WriteRecord(stream, 0x0207, payloads[0]);
            for (int i = 1; i < payloads.Count; i++) {
                WriteRecord(stream, 0x003c, payloads[i]);
            }
        }

        private sealed class LegacyXlsSharedStringTable {
            private readonly IReadOnlyList<LegacyXlsSharedStringEntry> _entries;
            private readonly Dictionary<long, uint> _indexesByCell;
            private readonly uint _totalCount;

            private LegacyXlsSharedStringTable(
                IReadOnlyList<LegacyXlsSharedStringEntry> entries,
                Dictionary<long, uint> indexesByCell,
                uint totalCount) {
                _entries = entries;
                _indexesByCell = indexesByCell;
                _totalCount = totalCount;
            }

            internal static LegacyXlsSharedStringTable Create(IReadOnlyList<List<LegacyXlsCell>> worksheets) {
                var entries = new List<LegacyXlsSharedStringEntry>();
                var indexesByValue = new Dictionary<LegacyXlsSharedStringKey, uint>();
                var indexesByCell = new Dictionary<long, uint>();
                uint totalCount = 0;

                for (int sheetIndex = 0; sheetIndex < worksheets.Count; sheetIndex++) {
                    foreach (LegacyXlsCell cell in worksheets[sheetIndex]) {
                        if (cell.Kind != LegacyXlsCellKind.Text) {
                            continue;
                        }

                        totalCount = checked(totalCount + 1U);
                        string text = cell.TextValue ?? string.Empty;
                        LegacyXlsTextFormattingRun[] formattingRuns = cell.TextFormattingRuns.ToArray();
                        var key = new LegacyXlsSharedStringKey(text, formattingRuns);
                        if (!indexesByValue.TryGetValue(key, out uint index)) {
                            index = checked((uint)entries.Count);
                            indexesByValue.Add(key, index);
                            entries.Add(new LegacyXlsSharedStringEntry(text, formattingRuns));
                        }

                        indexesByCell.Add(GetCellKey(sheetIndex, cell.Row, cell.Column), index);
                    }
                }

                return new LegacyXlsSharedStringTable(entries, indexesByCell, totalCount);
            }

            internal uint GetIndex(int sheetIndex, ushort row, ushort column) {
                if (_indexesByCell.TryGetValue(GetCellKey(sheetIndex, row, column), out uint index)) {
                    return index;
                }

                throw new InvalidOperationException("The native XLS shared-string table is missing a worksheet text cell.");
            }

            internal void WriteRecords(Stream stream) {
                if (_entries.Count == 0) {
                    return;
                }

                var writer = new LegacyXlsStringSegmentWriter();
                writer.WriteUInt32(_totalCount);
                writer.WriteUInt32(checked((uint)_entries.Count));
                var stringPositions = new List<LegacyXlsStringPosition>(_entries.Count);
                foreach (LegacyXlsSharedStringEntry entry in _entries) {
                    stringPositions.Add(writer.WriteUnicodeString(entry.Text, entry.FormattingRuns));
                }

                IReadOnlyList<byte[]> payloads = writer.GetPayloads();
                long sharedStringTableOffset = stream.Position;
                WriteRecord(stream, 0x00fc, payloads[0]);
                for (int i = 1; i < payloads.Count; i++) {
                    WriteRecord(stream, 0x003c, payloads[i]);
                }

                WriteExtSstRecord(stream, sharedStringTableOffset, payloads, stringPositions);
            }

            private void WriteExtSstRecord(
                Stream stream,
                long sharedStringTableOffset,
                IReadOnlyList<byte[]> sharedStringPayloads,
                IReadOnlyList<LegacyXlsStringPosition> stringPositions) {
                int stringsPerBucket = Math.Max((_entries.Count / 128) + 1, 8);
                long[] recordOffsets = new long[sharedStringPayloads.Count];
                long recordOffset = sharedStringTableOffset;
                for (int i = 0; i < sharedStringPayloads.Count; i++) {
                    recordOffsets[i] = recordOffset;
                    recordOffset = checked(recordOffset + 4L + sharedStringPayloads[i].Length);
                }

                using var payload = new MemoryStream();
                WriteUInt16(payload, checked((ushort)stringsPerBucket));
                for (int stringIndex = 0; stringIndex < stringPositions.Count; stringIndex += stringsPerBucket) {
                    LegacyXlsStringPosition position = stringPositions[stringIndex];
                    ushort recordRelativeOffset = checked((ushort)(position.Offset + 4));
                    uint containingRecordOffset = checked((uint)recordOffsets[position.SegmentIndex]);
                    WriteUInt32(payload, containingRecordOffset);
                    WriteUInt16(payload, recordRelativeOffset);
                    WriteUInt16(payload, 0);
                }

                WriteRecord(stream, 0x00ff, payload.ToArray());
            }

            private static long GetCellKey(int sheetIndex, ushort row, ushort column) {
                return ((long)sheetIndex << 32) | ((long)row << 16) | column;
            }
        }

        private sealed class LegacyXlsStringSegmentWriter {
            private readonly List<MemoryStream> _segments = new List<MemoryStream> { new MemoryStream() };

            private MemoryStream Current => _segments[_segments.Count - 1];

            internal void WriteUInt32(uint value) {
                EnsureAtomicCapacity(4);
                LegacyXlsWriter.WriteUInt32(Current, value);
            }

            internal LegacyXlsStringPosition WriteUnicodeString(string text, IReadOnlyList<LegacyXlsTextFormattingRun> formattingRuns) {
                if (text.Length > 32767) {
                    throw new NotSupportedException("Native XLS saving does not support strings longer than 32,767 characters.");
                }

                if (formattingRuns.Count > ushort.MaxValue) {
                    throw new NotSupportedException("Native XLS saving does not support rich-text run counts outside BIFF8 limits.");
                }

                bool compressed = CanUseCompressedString(text);
                byte options = compressed ? (byte)0 : (byte)1;
                if (formattingRuns.Count > 0) {
                    options |= 0x08;
                }

                int headerSize = formattingRuns.Count > 0 ? 5 : 3;
                EnsureAtomicCapacity(headerSize);
                var position = new LegacyXlsStringPosition(_segments.Count - 1, checked((int)Current.Length));
                LegacyXlsWriter.WriteUInt16(Current, checked((ushort)text.Length));
                Current.WriteByte(options);
                if (formattingRuns.Count > 0) {
                    LegacyXlsWriter.WriteUInt16(Current, checked((ushort)formattingRuns.Count));
                }

                WriteStringCharacters(text, compressed, (byte)(options & 0x01));
                foreach (LegacyXlsTextFormattingRun run in formattingRuns) {
                    EnsureAtomicCapacity(4);
                    LegacyXlsWriter.WriteUInt16(Current, run.StartCharacter);
                    LegacyXlsWriter.WriteUInt16(Current, run.FontIndex);
                }

                return position;
            }

            internal IReadOnlyList<byte[]> GetPayloads() {
                return _segments.Select(segment => segment.ToArray()).ToArray();
            }

            private void WriteStringCharacters(string text, bool compressed, byte continuationOptions) {
                byte[] bytes = compressed ? Encoding.ASCII.GetBytes(text) : Encoding.Unicode.GetBytes(text);
                int bytesPerCharacter = compressed ? 1 : 2;
                int sourceOffset = 0;
                while (sourceOffset < bytes.Length) {
                    int availableBytes = checked(BiffMaxRecordDataLength - (int)Current.Length);
                    int availableCharacters = availableBytes / bytesPerCharacter;
                    if (availableCharacters == 0) {
                        StartSegment();
                        Current.WriteByte(continuationOptions);
                        availableBytes = BiffMaxRecordDataLength - 1;
                        availableCharacters = availableBytes / bytesPerCharacter;
                    }

                    int remainingCharacters = (bytes.Length - sourceOffset) / bytesPerCharacter;
                    int charactersToWrite = Math.Min(remainingCharacters, availableCharacters);
                    int bytesToWrite = checked(charactersToWrite * bytesPerCharacter);
                    Current.Write(bytes, sourceOffset, bytesToWrite);
                    sourceOffset += bytesToWrite;
                }
            }

            private void EnsureAtomicCapacity(int byteCount) {
                if (Current.Length + byteCount > BiffMaxRecordDataLength) {
                    StartSegment();
                }
            }

            private void StartSegment() {
                if (Current.Length == 0) {
                    return;
                }

                _segments.Add(new MemoryStream());
            }
        }

        private sealed class LegacyXlsSharedStringEntry {
            internal LegacyXlsSharedStringEntry(string text, IReadOnlyList<LegacyXlsTextFormattingRun> formattingRuns) {
                Text = text;
                FormattingRuns = formattingRuns;
            }

            internal string Text { get; }

            internal IReadOnlyList<LegacyXlsTextFormattingRun> FormattingRuns { get; }
        }

        private readonly struct LegacyXlsStringPosition {
            internal LegacyXlsStringPosition(int segmentIndex, int offset) {
                SegmentIndex = segmentIndex;
                Offset = offset;
            }

            internal int SegmentIndex { get; }

            internal int Offset { get; }
        }

        private readonly struct LegacyXlsSharedStringKey : IEquatable<LegacyXlsSharedStringKey> {
            private readonly string _text;
            private readonly IReadOnlyList<LegacyXlsTextFormattingRun> _formattingRuns;

            internal LegacyXlsSharedStringKey(string text, IReadOnlyList<LegacyXlsTextFormattingRun> formattingRuns) {
                _text = text;
                _formattingRuns = formattingRuns;
            }

            public bool Equals(LegacyXlsSharedStringKey other) {
                if (!string.Equals(_text, other._text, StringComparison.Ordinal)
                    || _formattingRuns.Count != other._formattingRuns.Count) {
                    return false;
                }

                for (int i = 0; i < _formattingRuns.Count; i++) {
                    LegacyXlsTextFormattingRun left = _formattingRuns[i];
                    LegacyXlsTextFormattingRun right = other._formattingRuns[i];
                    if (left.StartCharacter != right.StartCharacter || left.FontIndex != right.FontIndex) {
                        return false;
                    }
                }

                return true;
            }

            public override bool Equals(object? obj) {
                return obj is LegacyXlsSharedStringKey other && Equals(other);
            }

            public override int GetHashCode() {
                unchecked {
                    int hash = StringComparer.Ordinal.GetHashCode(_text);
                    for (int i = 0; i < _formattingRuns.Count; i++) {
                        hash = (hash * 397) ^ _formattingRuns[i].StartCharacter;
                        hash = (hash * 397) ^ _formattingRuns[i].FontIndex;
                    }

                    return hash;
                }
            }
        }
    }
}

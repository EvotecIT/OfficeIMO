using OfficeIMO.Excel.LegacyXls.Model;
using System.Text;
using System.Threading;

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
            private readonly uint _totalCount;
            private readonly Dictionary<string, uint>? _directIndexesByValue;

            internal LegacyXlsSharedStringTable(
                IReadOnlyList<LegacyXlsSharedStringEntry> entries,
                uint totalCount,
                Dictionary<string, uint>? directIndexesByValue = null) {
                _entries = entries;
                _totalCount = totalCount;
                _directIndexesByValue = directIndexesByValue;
            }

            internal static LegacyXlsSharedStringTable Create(IReadOnlyList<List<LegacyXlsCell>> worksheets) {
                var entries = new List<LegacyXlsSharedStringEntry>();
                var indexesByValue = new Dictionary<LegacyXlsSharedStringKey, uint>();
                uint totalCount = 0;

                for (int sheetIndex = 0; sheetIndex < worksheets.Count; sheetIndex++) {
                    List<LegacyXlsCell> worksheet = worksheets[sheetIndex];
                    for (int cellIndex = 0; cellIndex < worksheet.Count; cellIndex++) {
                        LegacyXlsCell cell = worksheet[cellIndex];
                        if (cell.Kind != LegacyXlsCellKind.Text) {
                            continue;
                        }

                        totalCount = checked(totalCount + 1U);
                        string text = cell.TextValue ?? string.Empty;
                        LegacyXlsTextFormattingRun[] formattingRuns = cell.TextFormattingRuns.Count == 0
                            ? Array.Empty<LegacyXlsTextFormattingRun>()
                            : cell.TextFormattingRuns.ToArray();
                        var key = new LegacyXlsSharedStringKey(text, formattingRuns);
                        if (!indexesByValue.TryGetValue(key, out uint index)) {
                            index = checked((uint)entries.Count);
                            indexesByValue.Add(key, index);
                            entries.Add(new LegacyXlsSharedStringEntry(text, formattingRuns));
                        }

                        worksheet[cellIndex] = cell.WithSharedStringIndex(index);
                    }
                }

                return new LegacyXlsSharedStringTable(entries, totalCount);
            }

            internal uint GetIndex(LegacyXlsCell cell) {
                if (cell.SharedStringIndex != uint.MaxValue) {
                    return cell.SharedStringIndex;
                }

                throw new InvalidOperationException("The native XLS shared-string table is missing a worksheet text cell.");
            }

            internal uint GetDirectIndex(string text) {
                if (_directIndexesByValue != null
                    && _directIndexesByValue.TryGetValue(text, out uint index)) {
                    return index;
                }

                throw new InvalidOperationException("The native XLS shared-string table is missing a direct worksheet text cell.");
            }

            internal bool TryGetDirectIndex(string text, out uint index) {
                if (_directIndexesByValue != null) {
                    return _directIndexesByValue.TryGetValue(text, out index);
                }

                index = 0;
                return false;
            }

            internal void WriteRecords(Stream stream, CancellationToken cancellationToken = default, long maximumBytes = int.MaxValue) {
                cancellationToken.ThrowIfCancellationRequested();
                if (_entries.Count == 0) {
                    return;
                }

                var writer = new LegacyXlsStringSegmentWriter(maximumBytes);
                writer.WriteUInt32(_totalCount);
                writer.WriteUInt32(checked((uint)_entries.Count));
                var stringPositions = new List<LegacyXlsStringPosition>(_entries.Count);
                for (int entryIndex = 0; entryIndex < _entries.Count; entryIndex++) {
                    if ((entryIndex & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                    LegacyXlsSharedStringEntry entry = _entries[entryIndex];
                    stringPositions.Add(writer.WriteUnicodeString(entry.Text, entry.FormattingRuns));
                }

                cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyList<byte[]> payloads = writer.GetPayloads(cancellationToken);
                long sharedStringTableOffset = stream.Position;
                WriteRecord(stream, 0x00fc, payloads[0]);
                for (int i = 1; i < payloads.Count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    WriteRecord(stream, 0x003c, payloads[i]);
                }

                WriteExtSstRecord(
                    stream,
                    sharedStringTableOffset,
                    payloads,
                    stringPositions,
                    cancellationToken);
            }

            private void WriteExtSstRecord(
                Stream stream,
                long sharedStringTableOffset,
                IReadOnlyList<byte[]> sharedStringPayloads,
                IReadOnlyList<LegacyXlsStringPosition> stringPositions,
                CancellationToken cancellationToken) {
                int stringsPerBucket = Math.Max((_entries.Count / 128) + 1, 8);
                long[] recordOffsets = new long[sharedStringPayloads.Count];
                long recordOffset = sharedStringTableOffset;
                for (int i = 0; i < sharedStringPayloads.Count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    recordOffsets[i] = recordOffset;
                    recordOffset = checked(recordOffset + 4L + sharedStringPayloads[i].Length);
                }

                using var payload = new MemoryStream();
                WriteUInt16(payload, checked((ushort)stringsPerBucket));
                for (int stringIndex = 0; stringIndex < stringPositions.Count; stringIndex += stringsPerBucket) {
                    cancellationToken.ThrowIfCancellationRequested();
                    LegacyXlsStringPosition position = stringPositions[stringIndex];
                    ushort recordRelativeOffset = checked((ushort)(position.Offset + 4));
                    uint containingRecordOffset = checked((uint)recordOffsets[position.SegmentIndex]);
                    WriteUInt32(payload, checked(containingRecordOffset + recordRelativeOffset));
                    WriteUInt16(payload, recordRelativeOffset);
                    WriteUInt16(payload, 0);
                }

                cancellationToken.ThrowIfCancellationRequested();
                WriteRecord(stream, 0x00ff, payload.ToArray());
            }
        }

        private sealed class LegacyXlsDirectSharedStringBuilder {
            private readonly List<LegacyXlsSharedStringEntry> _entries = new();
            private readonly Dictionary<string, uint> _indexesByValue = new(StringComparer.Ordinal);
            private uint _totalCount;
            private long _estimatedSerializedBytes;

            internal long EstimatedSerializedBytes => _estimatedSerializedBytes;

            internal uint Add(string text) {
                _totalCount = checked(_totalCount + 1U);
                if (_indexesByValue.TryGetValue(text, out uint existingIndex)) {
                    return existingIndex;
                }

                uint index = checked((uint)_entries.Count);
                _indexesByValue.Add(text, index);
                _entries.Add(new LegacyXlsSharedStringEntry(
                    text,
                    Array.Empty<LegacyXlsTextFormattingRun>()));
                _estimatedSerializedBytes = checked(
                    _estimatedSerializedBytes
                    + 3L
                    + (text.Length * (CanUseCompressedString(text) ? 1L : 2L)));
                return index;
            }

            internal LegacyXlsSharedStringTable Build() =>
                new LegacyXlsSharedStringTable(_entries, _totalCount, _indexesByValue);
        }

        private sealed class LegacyXlsStringSegmentWriter {
            // Each segment is one bounded record payload. Owned buffers avoid a stream
            // call and limit check per header byte while keeping the same bounds.
            private readonly List<byte[]> _segments = new();
            private readonly List<int> _segmentLengths = new();
            private readonly long _maximumBytes;
            private long _completedBytes;
            private byte[] _current = null!;
            private int _currentLength;
            private int _currentLimit;
            private byte[] _scratch = new byte[64];

            internal LegacyXlsStringSegmentWriter(long maximumBytes = int.MaxValue) {
                _maximumBytes = maximumBytes;
                AddBoundedSegment();
            }

            private void AddBoundedSegment() {
                long remaining = _maximumBytes - _completedBytes - 4;
                if (remaining <= 0) throw OfficeIMO.Core.Internal.OfficeOutputLimit.Create("Native XLS shared strings exceed the in-memory save limit.");
                if (_segments.Count > 0) _segmentLengths[_segments.Count - 1] = _currentLength;
                _currentLimit = (int)Math.Min(BiffMaxRecordDataLength, remaining);
                _current = new byte[Math.Min(256, _currentLimit)];
                _currentLength = 0;
                _segments.Add(_current);
                _segmentLengths.Add(0);
            }

            internal void WriteUInt32(uint value) {
                EnsureAtomicCapacity(4);
                byte[] bytes = Reserve(4);
                bytes[_currentLength] = (byte)value;
                bytes[_currentLength + 1] = (byte)(value >> 8);
                bytes[_currentLength + 2] = (byte)(value >> 16);
                bytes[_currentLength + 3] = (byte)(value >> 24);
                _currentLength += 4;
            }

            internal LegacyXlsStringPosition WriteUnicodeString(string text, IReadOnlyList<LegacyXlsTextFormattingRun> formattingRuns) {
                if (text.Length > 32767) {
                    throw new NotSupportedException("Native XLS saving does not support strings longer than 32,767 characters.");
                }

                int runCount = formattingRuns.Count;
                if (runCount > ushort.MaxValue) {
                    throw new NotSupportedException("Native XLS saving does not support rich-text run counts outside BIFF8 limits.");
                }

                bool compressed = CanUseCompressedString(text);
                byte options = compressed ? (byte)0 : (byte)1;
                if (runCount > 0) {
                    options |= 0x08;
                }

                int headerSize = runCount > 0 ? 5 : 3;
                EnsureAtomicCapacity(headerSize);
                var position = new LegacyXlsStringPosition(_segments.Count - 1, _currentLength);
                byte[] header = Reserve(headerSize);
                header[_currentLength] = (byte)text.Length;
                header[_currentLength + 1] = (byte)(text.Length >> 8);
                header[_currentLength + 2] = options;
                if (runCount > 0) {
                    header[_currentLength + 3] = (byte)runCount;
                    header[_currentLength + 4] = (byte)(runCount >> 8);
                }
                _currentLength += headerSize;

                WriteStringCharacters(text, compressed, (byte)(options & 0x01));
                for (int index = 0; index < runCount; index++) {
                    LegacyXlsTextFormattingRun run = formattingRuns[index];
                    EnsureAtomicCapacity(4);
                    byte[] bytes = Reserve(4);
                    bytes[_currentLength] = (byte)run.StartCharacter;
                    bytes[_currentLength + 1] = (byte)(run.StartCharacter >> 8);
                    bytes[_currentLength + 2] = (byte)run.FontIndex;
                    bytes[_currentLength + 3] = (byte)(run.FontIndex >> 8);
                    _currentLength += 4;
                }

                return position;
            }

            internal IReadOnlyList<byte[]> GetPayloads(CancellationToken cancellationToken = default) {
                _segmentLengths[_segments.Count - 1] = _currentLength;
                var payloads = new byte[_segments.Count][];
                for (int segmentIndex = 0; segmentIndex < _segments.Count; segmentIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int length = _segmentLengths[segmentIndex];
                    var payload = new byte[length];
                    Buffer.BlockCopy(_segments[segmentIndex], 0, payload, 0, length);
                    payloads[segmentIndex] = payload;
                }

                cancellationToken.ThrowIfCancellationRequested();
                return payloads;
            }

            private void WriteStringCharacters(string text, bool compressed, byte continuationOptions) {
                int bytesPerCharacter = compressed ? 1 : 2;
                int byteCount = text.Length * bytesPerCharacter;
                if (_scratch.Length < byteCount) _scratch = new byte[Math.Max(byteCount, _scratch.Length * 2)];
                byte[] source = _scratch;
                // Compressed text is 7-bit ASCII; both encoders keep their existing byte output.
                _ = compressed
                    ? Encoding.ASCII.GetBytes(text, 0, text.Length, source, 0)
                    : Encoding.Unicode.GetBytes(text, 0, text.Length, source, 0);
                int sourceOffset = 0;
                while (sourceOffset < byteCount) {
                    int availableBytes = BiffMaxRecordDataLength - _currentLength;
                    int availableCharacters = availableBytes / bytesPerCharacter;
                    if (availableCharacters == 0) {
                        StartSegment();
                        Reserve(1)[_currentLength++] = continuationOptions;
                        availableBytes = BiffMaxRecordDataLength - 1;
                        availableCharacters = availableBytes / bytesPerCharacter;
                    }

                    int remainingCharacters = (byteCount - sourceOffset) / bytesPerCharacter;
                    int charactersToWrite = Math.Min(remainingCharacters, availableCharacters);
                    int bytesToWrite = checked(charactersToWrite * bytesPerCharacter);
                    Buffer.BlockCopy(source, sourceOffset, Reserve(bytesToWrite), _currentLength, bytesToWrite);
                    _currentLength += bytesToWrite;
                    sourceOffset += bytesToWrite;
                }
            }

            private byte[] Reserve(int byteCount) {
                int end = _currentLength + byteCount;
                if (end > _currentLimit) {
                    throw OfficeIMO.Core.Internal.OfficeOutputLimit.Create($"The artifact exceeds the configured output limit of {_currentLimit} bytes.");
                }
                if (end > _current.Length) {
                    var grown = new byte[Math.Min(_currentLimit, Math.Max(end, _current.Length * 2))];
                    Buffer.BlockCopy(_current, 0, grown, 0, _currentLength);
                    _current = grown;
                    _segments[_segments.Count - 1] = grown;
                }
                return _current;
            }

            private void EnsureAtomicCapacity(int byteCount) {
                if (_currentLength + byteCount > BiffMaxRecordDataLength) {
                    StartSegment();
                }
            }

            private void StartSegment() {
                if (_currentLength == 0) {
                    return;
                }

                _completedBytes = checked(_completedBytes + _currentLength + 4);
                AddBoundedSegment();
            }
        }

        // A value entry keeps one table slot per distinct string without an object per string.
        private readonly struct LegacyXlsSharedStringEntry {
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

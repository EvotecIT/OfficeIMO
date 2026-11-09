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
            private readonly uint _totalCount;
            private readonly int _entryCount;
            private readonly Dictionary<string, uint>? _directIndexesByValue;

            /// <summary>Describes the active entry prefix independently of retained workspace capacity.</summary>
            internal LegacyXlsSharedStringTable(
                IReadOnlyList<LegacyXlsSharedStringEntry> entries,
                uint totalCount,
                Dictionary<string, uint>? directIndexesByValue = null,
                int? entryCount = null) {
                _entries = entries;
                _entryCount = entryCount ?? entries.Count;
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

            /// <summary>Writes the active shared strings and their extension indexes within the output limit.</summary>
            internal void WriteRecords(Stream stream, CancellationToken cancellationToken = default, long maximumBytes = int.MaxValue) {
                cancellationToken.ThrowIfCancellationRequested();
                if (_entryCount == 0) {
                    return;
                }

                // Three header bytes per string provide a conservative capacity bound.
                int segmentCapacity = (int)Math.Min(BiffMaxRecordDataLength, Math.Max(256L, 8L + 3L * _entryCount));
                long sharedStringTableOffset = stream.Position;
                var writer = new LegacyXlsStringSegmentWriter(maximumBytes, segmentCapacity, stream, cancellationToken);
                writer.WriteUInt32(_totalCount);
                writer.WriteUInt32(checked((uint)_entryCount));
                int stringsPerBucket = Math.Max((_entryCount / 128) + 1, 8);
                var stringPositions = new List<LegacyXlsStringPosition>(((_entryCount - 1) / stringsPerBucket) + 1);
                int nextBucketIndex = 0;
                for (int entryIndex = 0; entryIndex < _entryCount; entryIndex++) {
                    if ((entryIndex & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                    LegacyXlsSharedStringEntry entry = _entries[entryIndex];
                    LegacyXlsStringPosition position = writer.WriteUnicodeString(entry.Text, entry.FormattingRuns);
                    if (entryIndex == nextBucketIndex) {
                        stringPositions.Add(position);
                        nextBucketIndex += stringsPerBucket;
                    }
                }

                cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyList<int> payloadLengths = writer.WriteSharedStringRecords(stream, cancellationToken);

                WriteExtSstRecord(
                    stream,
                    sharedStringTableOffset,
                    payloadLengths,
                    stringPositions,
                    stringsPerBucket,
                    cancellationToken);
            }

            private void WriteExtSstRecord(
                Stream stream,
                long sharedStringTableOffset,
                IReadOnlyList<int> sharedStringPayloadLengths,
                IReadOnlyList<LegacyXlsStringPosition> stringPositions,
                int stringsPerBucket,
                CancellationToken cancellationToken) {
                long[] recordOffsets = new long[sharedStringPayloadLengths.Count];
                long recordOffset = sharedStringTableOffset;
                for (int i = 0; i < sharedStringPayloadLengths.Count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    recordOffsets[i] = recordOffset;
                    recordOffset = checked(recordOffset + 4L + sharedStringPayloadLengths[i]);
                }

                using var payload = new MemoryStream();
                WriteUInt16(payload, checked((ushort)stringsPerBucket));
                for (int stringIndex = 0; stringIndex < stringPositions.Count; stringIndex++) {
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
            private const int MaximumReusableStrings = 65_536;
            private readonly Dictionary<string, uint> _indexesByValue = new(StringComparer.Ordinal);
            private LegacyXlsSharedStringEntry[]? _reusableEntries;
            private uint _totalCount;
            private long _estimatedSerializedBytes;

            internal long EstimatedSerializedBytes => _estimatedSerializedBytes;

            internal uint Add(string text) {
                _totalCount = checked(_totalCount + 1U);
#if NET8_0_OR_GREATER
                ref uint index = ref System.Runtime.InteropServices.CollectionsMarshal.GetValueRefOrAddDefault(
                    _indexesByValue, text, out bool exists);
                if (exists) return index;
                index = checked((uint)(_indexesByValue.Count - 1));
#else
                if (_indexesByValue.TryGetValue(text, out uint existingIndex)) return existingIndex;
                uint index = checked((uint)_indexesByValue.Count);
                _indexesByValue.Add(text, index);
#endif
                _estimatedSerializedBytes = checked(
                    _estimatedSerializedBytes
                    + 3L
                    + (text.Length * (CanUseCompressedString(text) ? 1L : 2L)));
                return index;
            }

            /// <summary>Builds an insertion-indexed table, optionally borrowing this exclusive workspace's entry buffer.</summary>
            internal LegacyXlsSharedStringTable Build(bool reuseEntries = false) {
                int count = _indexesByValue.Count;
                LegacyXlsSharedStringEntry[] entries;
                if (count == 0) {
                    entries = Array.Empty<LegacyXlsSharedStringEntry>();
                } else if (reuseEntries && count <= MaximumReusableStrings) {
                    if (_reusableEntries == null || _reusableEntries.Length < count) {
                        _reusableEntries = new LegacyXlsSharedStringEntry[count];
                    }
                    entries = _reusableEntries;
                } else {
                    entries = new LegacyXlsSharedStringEntry[count];
                }
                // The recorded indexes keep the table independent of dictionary enumeration order.
                foreach (KeyValuePair<string, uint> entry in _indexesByValue) {
                    entries[checked((int)entry.Value)] = new LegacyXlsSharedStringEntry(
                        entry.Key, Array.Empty<LegacyXlsTextFormattingRun>());
                }
                return new LegacyXlsSharedStringTable(entries, _totalCount, _indexesByValue, count);
            }

            /// <summary>Releases all text references and counters before an exclusive workspace is reused.</summary>
            internal bool ClearForReuse() {
                bool retain = _indexesByValue.Count <= MaximumReusableStrings;
                _indexesByValue.Clear();
                _totalCount = 0;
                _estimatedSerializedBytes = 0;
                if (_reusableEntries != null) {
                    Array.Clear(_reusableEntries, 0, _reusableEntries.Length);
                }
                return retain;
            }
        }

        private sealed class LegacyXlsStringSegmentWriter {
            // Preservation retains bounded payloads. Workbook saves write completed
            // records into their internal stream and reuse the active buffer.
            private readonly List<byte[]> _segments = new();
            private readonly List<int> _segmentLengths = new();
            private readonly long _maximumBytes;
            private readonly Stream? _output;
            private readonly CancellationToken _cancellationToken;
            private readonly int _initialSegmentCapacity;
            private long _completedBytes;
            private byte[] _current = null!;
            private int _currentLength;
            private int _currentLimit;
            private byte[] _scratch = new byte[64];

            /// <summary>Buffers preservation payloads or reuses a payload buffer for an internal workbook stream.</summary>
            internal LegacyXlsStringSegmentWriter(long maximumBytes = int.MaxValue, int initialSegmentCapacity = 256,
                Stream? output = null, CancellationToken cancellationToken = default) {
                _maximumBytes = maximumBytes;
                _output = output;
                _cancellationToken = cancellationToken;
                _initialSegmentCapacity = initialSegmentCapacity;
                AddBoundedSegment();
            }

            /// <summary>Starts a bounded payload after delivering its completed predecessor when output is supplied.</summary>
            private void AddBoundedSegment() {
                long remaining = _maximumBytes - _completedBytes - 4;
                if (remaining <= 0) throw OfficeIMO.Core.Internal.OfficeOutputLimit.Create("Native XLS shared strings exceed the in-memory save limit.");
                if (_segments.Count > 0) {
                    _segmentLengths[_segments.Count - 1] = _currentLength;
                    if (_output != null) {
                        _cancellationToken.ThrowIfCancellationRequested();
                        WriteRecord(_output, _segments.Count == 1 ? (ushort)0x00fc : (ushort)0x003c,
                            _current, _currentLength);
                    }
                }
                _currentLimit = (int)Math.Min(BiffMaxRecordDataLength, remaining);
                int capacity = Math.Min(_initialSegmentCapacity, _currentLimit);
                if (_output == null || _current == null || _current.Length < capacity) {
                    _current = new byte[capacity];
                }
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

            /// <summary>Finishes the SST and continuation records and returns their used payload lengths.</summary>
            internal IReadOnlyList<int> WriteSharedStringRecords(Stream stream, CancellationToken cancellationToken) {
                _segmentLengths[_segments.Count - 1] = _currentLength;
                if (_output != null) {
                    cancellationToken.ThrowIfCancellationRequested();
                    WriteRecord(_output, _segments.Count == 1 ? (ushort)0x00fc : (ushort)0x003c,
                        _current, _currentLength);
                    cancellationToken.ThrowIfCancellationRequested();
                    return _segmentLengths;
                }
                for (int index = 0; index < _segments.Count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    WriteRecord(stream, index == 0 ? (ushort)0x00fc : (ushort)0x003c,
                        _segments[index], _segmentLengths[index]);
                }
                cancellationToken.ThrowIfCancellationRequested();
                return _segmentLengths;
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

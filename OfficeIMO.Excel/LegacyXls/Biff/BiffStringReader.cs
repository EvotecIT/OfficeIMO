using OfficeIMO.Excel.LegacyXls.Diagnostics;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffStringReader {
        internal static string ReadShortUnicodeString(byte[] bytes, ref int offset) {
            if (offset >= bytes.Length) throw new InvalidDataException("Unexpected end of BIFF short string.");
            int charCount = bytes[offset++];
            return ReadUnicodeStringBody(bytes, ref offset, charCount);
        }

        internal static string ReadUnicodeString(byte[] bytes, ref int offset) {
            int charCount = BiffRecordReader.ReadUInt16(bytes, offset);
            offset += 2;
            return ReadUnicodeStringBody(bytes, ref offset, charCount);
        }

        internal static string ReadUnicodeString(IReadOnlyList<byte[]> payloads) {
            if (payloads.Count == 0) {
                throw new InvalidDataException("The BIFF string has no payload data.");
            }

            var reader = new BiffStringSegmentReader(payloads);
            return ReadSegmentedUnicodeString(reader);
        }

        internal static string ReadUnicodeStringNoCch(byte[] bytes, ref int offset, int charCount) {
            return ReadUnicodeStringBody(bytes, ref offset, charCount);
        }

        internal static string ReadWideString(byte[] bytes, ref int offset) {
            if (offset + 2 > bytes.Length) {
                throw new InvalidDataException("Unexpected end of BIFF wide string length.");
            }

            int charCount = BiffRecordReader.ReadUInt16(bytes, offset);
            offset += 2;
            int byteCount = checked(charCount * 2);
            if (offset + byteCount > bytes.Length) {
                throw new InvalidDataException("Unexpected end of BIFF wide string characters.");
            }

            string value = Encoding.Unicode.GetString(bytes, offset, byteCount);
            offset += byteCount;
            return value;
        }

        internal static List<string> ReadSharedStrings(byte[] payload, List<LegacyXlsImportDiagnostic> diagnostics, int recordOffset) {
            return ReadSharedStrings(new[] { payload }, diagnostics, recordOffset);
        }

        internal static List<string> ReadSharedStrings(IReadOnlyList<byte[]> payloads, List<LegacyXlsImportDiagnostic> diagnostics, int recordOffset) {
            var strings = new List<string>();
            if (payloads.Count == 0 || payloads[0].Length < 8) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-SST-SHORT",
                    "The shared string table record is too short.",
                    recordOffset: recordOffset,
                    recordType: (ushort)BiffRecordType.Sst));
                return strings;
            }

            var reader = new BiffStringSegmentReader(payloads);
            _ = reader.ReadUInt32Raw();
            uint uniqueCount = reader.ReadUInt32Raw();
            for (uint i = 0; i < uniqueCount && reader.HasData; i++) {
                try {
                    strings.Add(ReadSegmentedUnicodeString(reader));
                } catch (InvalidDataException ex) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-SST-STRING-INVALID",
                        FormattableString.Invariant($"Shared string index {i} could not be read. {ex.Message}"),
                        recordOffset: recordOffset,
                        recordType: (ushort)BiffRecordType.Sst));
                    break;
                }
            }

            return strings;
        }

        private static string ReadSegmentedUnicodeString(BiffStringSegmentReader reader) {
            int charCount = reader.ReadUInt16Raw();
            return ReadSegmentedUnicodeStringBody(reader, charCount);
        }

        private static string ReadUnicodeStringBody(byte[] bytes, ref int offset, int charCount) {
            if (offset >= bytes.Length) throw new InvalidDataException("Unexpected end of BIFF string options.");
            byte options = bytes[offset++];
            bool isUtf16 = (options & 0x01) != 0;
            bool hasExtended = (options & 0x04) != 0;
            bool hasRichText = (options & 0x08) != 0;
            ushort richTextRuns = 0;
            uint extendedSize = 0;

            if (hasRichText) {
                richTextRuns = BiffRecordReader.ReadUInt16(bytes, offset);
                offset += 2;
            }

            if (hasExtended) {
                extendedSize = BiffRecordReader.ReadUInt32(bytes, offset);
                offset += 4;
            }

            int byteCount = checked(charCount * (isUtf16 ? 2 : 1));
            if (offset + byteCount > bytes.Length) {
                throw new InvalidDataException("Unexpected end of BIFF string characters.");
            }

            string value = isUtf16
                ? Encoding.Unicode.GetString(bytes, offset, byteCount)
                : ReadCompressedUnicode(bytes, offset, byteCount);
            offset += byteCount;

            int formattingBytes = checked(richTextRuns * 4);
            int extraBytes = checked(formattingBytes + (int)extendedSize);
            if (offset + extraBytes > bytes.Length) {
                throw new InvalidDataException("Unexpected end of BIFF string formatting data.");
            }

            offset += extraBytes;
            return value;
        }

        private static string ReadSegmentedUnicodeStringBody(BiffStringSegmentReader reader, int charCount) {
            byte options = reader.ReadByteRaw();
            bool isUtf16 = (options & 0x01) != 0;
            bool hasExtended = (options & 0x04) != 0;
            bool hasRichText = (options & 0x08) != 0;
            ushort richTextRuns = 0;
            uint extendedSize = 0;

            if (hasRichText) {
                richTextRuns = reader.ReadUInt16Raw();
            }

            if (hasExtended) {
                extendedSize = reader.ReadUInt32Raw();
            }

            string value = reader.ReadStringCharacters(charCount, isUtf16);
            int formattingBytes = checked(richTextRuns * 4);
            int extraBytes = checked(formattingBytes + (int)extendedSize);
            reader.SkipStringVariableBytes(extraBytes);
            return value;
        }

        private static string ReadCompressedUnicode(byte[] bytes, int offset, int byteCount) {
            char[] chars = new char[byteCount];
            for (int i = 0; i < byteCount; i++) {
                chars[i] = (char)bytes[offset + i];
            }

            return new string(chars);
        }

        private sealed class BiffStringSegmentReader {
            private readonly IReadOnlyList<byte[]> _segments;
            private int _segmentIndex;
            private int _offset;

            internal BiffStringSegmentReader(IReadOnlyList<byte[]> segments) {
                _segments = segments ?? throw new ArgumentNullException(nameof(segments));
            }

            internal bool HasData {
                get {
                    AdvancePastEmptySegments();
                    return _segmentIndex < _segments.Count && _offset < _segments[_segmentIndex].Length;
                }
            }

            internal byte ReadByteRaw() {
                AdvancePastEmptySegments();
                if (_segmentIndex >= _segments.Count) {
                    throw new InvalidDataException("Unexpected end of BIFF string data.");
                }

                return _segments[_segmentIndex][_offset++];
            }

            internal ushort ReadUInt16Raw() {
                int low = ReadByteRaw();
                int high = ReadByteRaw();
                return (ushort)(low | (high << 8));
            }

            internal uint ReadUInt32Raw() {
                uint b0 = ReadByteRaw();
                uint b1 = ReadByteRaw();
                uint b2 = ReadByteRaw();
                uint b3 = ReadByteRaw();
                return b0 | (b1 << 8) | (b2 << 16) | (b3 << 24);
            }

            internal string ReadStringCharacters(int charCount, bool isUtf16) {
                if (charCount < 0) {
                    throw new InvalidDataException("The BIFF string character count is invalid.");
                }

                var builder = new StringBuilder(charCount);
                bool currentUtf16 = isUtf16;
                int remaining = charCount;
                while (remaining > 0) {
                    EnsureStringVariableDataAvailable(ref currentUtf16);
                    byte[] segment = _segments[_segmentIndex];
                    int availableBytes = segment.Length - _offset;
                    if (currentUtf16) {
                        int availableChars = availableBytes / 2;
                        if (availableChars <= 0) {
                            throw new InvalidDataException("A continued BIFF Unicode string split a double-byte character.");
                        }

                        int take = Math.Min(remaining, availableChars);
                        builder.Append(Encoding.Unicode.GetString(segment, _offset, checked(take * 2)));
                        _offset += checked(take * 2);
                        remaining -= take;
                    } else {
                        int take = Math.Min(remaining, availableBytes);
                        for (int i = 0; i < take; i++) {
                            builder.Append((char)segment[_offset + i]);
                        }

                        _offset += take;
                        remaining -= take;
                    }
                }

                return builder.ToString();
            }

            internal void SkipStringVariableBytes(int byteCount) {
                if (byteCount < 0) {
                    throw new InvalidDataException("The BIFF string variable data length is invalid.");
                }

                int remaining = byteCount;
                while (remaining > 0) {
                    EnsureRawDataAvailable();
                    byte[] segment = _segments[_segmentIndex];
                    int take = Math.Min(remaining, segment.Length - _offset);
                    _offset += take;
                    remaining -= take;
                }
            }

            private void EnsureStringVariableDataAvailable(ref bool isUtf16) {
                if (_segmentIndex >= _segments.Count) {
                    throw new InvalidDataException("Unexpected end of BIFF string variable data.");
                }

                if (_offset < _segments[_segmentIndex].Length) {
                    return;
                }

                _segmentIndex++;
                _offset = 0;
                if (_segmentIndex >= _segments.Count || _segments[_segmentIndex].Length == 0) {
                    throw new InvalidDataException("A BIFF string continuation is missing its character option byte.");
                }

                byte continueOptions = _segments[_segmentIndex][_offset++];
                isUtf16 = (continueOptions & 0x01) != 0;
            }

            private void EnsureRawDataAvailable() {
                if (_segmentIndex >= _segments.Count) {
                    throw new InvalidDataException("Unexpected end of BIFF string variable data.");
                }

                if (_offset < _segments[_segmentIndex].Length) {
                    return;
                }

                _segmentIndex++;
                _offset = 0;
                if (_segmentIndex >= _segments.Count || _segments[_segmentIndex].Length == 0) {
                    throw new InvalidDataException("Unexpected end of BIFF string variable data.");
                }
            }

            private void AdvancePastEmptySegments() {
                while (_segmentIndex < _segments.Count
                    && _offset >= _segments[_segmentIndex].Length
                    && _segmentIndex + 1 < _segments.Count) {
                    _segmentIndex++;
                    _offset = 0;
                }
            }
        }
    }
}

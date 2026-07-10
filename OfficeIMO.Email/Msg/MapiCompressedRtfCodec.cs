namespace OfficeIMO.Email;

/// <summary>
/// Implements the MS-OXRTFCP transport encoding used by PidTagRtfCompressed.
/// RTF syntax and semantic conversion remain owned by OfficeIMO.Rtf.
/// </summary>
internal static class MapiCompressedRtfCodec {
    private const int HeaderLength = 16;
    private const uint CompressedType = 0x75465A4C;
    private const uint UncompressedType = 0x414C454D;

    internal static bool TryDecompress(byte[] source, long maximumOutputBytes,
        IList<EmailDiagnostic> diagnostics, string location, CancellationToken cancellationToken,
        out byte[] result) {
        result = Array.Empty<byte>();
        cancellationToken.ThrowIfCancellationRequested();
        if (source.Length < HeaderLength) {
            AddError(diagnostics, "EMAIL_MSG_RTF_HEADER_TRUNCATED",
                "The compressed-RTF header is truncated.", location);
            return false;
        }

        uint compressedSize = MsgBinary.ReadUInt32(source, 0);
        uint rawSize = MsgBinary.ReadUInt32(source, 4);
        uint compressionType = MsgBinary.ReadUInt32(source, 8);
        uint expectedCrc = MsgBinary.ReadUInt32(source, 12);
        if (compressedSize < 12 || compressedSize - 12 > int.MaxValue) {
            AddError(diagnostics, "EMAIL_MSG_RTF_SIZE_INVALID",
                "The compressed-RTF content size is invalid.", location);
            return false;
        }
        if (rawSize > int.MaxValue || rawSize > maximumOutputBytes) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxDecodedPropertyBytes),
                rawSize, maximumOutputBytes);
        }

        int contentsLength = checked((int)(compressedSize - 12));
        if (contentsLength > source.Length - HeaderLength) {
            AddError(diagnostics, "EMAIL_MSG_RTF_SIZE_INVALID",
                "The compressed-RTF content extends beyond the property stream.", location);
            return false;
        }

        if (compressionType == UncompressedType) {
            if (rawSize > contentsLength) {
                AddError(diagnostics, "EMAIL_MSG_RTF_SIZE_MISMATCH",
                    "The uncompressed RTF content is shorter than its declared raw size.", location);
                return false;
            }
            if (rawSize != contentsLength) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_RTF_SIZE_MISMATCH",
                    "The uncompressed RTF content contains bytes beyond its declared raw size; they were ignored.",
                    EmailDiagnosticSeverity.Warning, location));
            }
            result = MsgBinary.Slice(source, HeaderLength, checked((int)rawSize));
            return true;
        }

        if (compressionType != CompressedType) {
            AddError(diagnostics, "EMAIL_MSG_RTF_COMPRESSION_UNKNOWN",
                "The PidTagRtfCompressed property uses an unknown compression type.", location);
            return false;
        }

        uint actualCrc = CalculateCrc(source, HeaderLength, contentsLength);
        if (actualCrc != expectedCrc) {
            AddError(diagnostics, "EMAIL_MSG_RTF_CRC_MISMATCH",
                "The compressed-RTF checksum does not match the content.", location);
            return false;
        }

        var dictionary = new CompressionDictionary();
        using (var output = new MemoryStream(checked((int)rawSize))) {
            int inputOffset = HeaderLength;
            int inputEnd = HeaderLength + contentsLength;
            bool terminated = false;
            bool ignoredEmptySentinel = false;
            while (inputOffset < inputEnd && !terminated) {
                cancellationToken.ThrowIfCancellationRequested();
                byte control = source[inputOffset++];
                for (int token = 0; token < 8; token++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if ((control & (1 << token)) == 0) {
                        if (inputOffset >= inputEnd) {
                            AddError(diagnostics, "EMAIL_MSG_RTF_CONTENT_TRUNCATED",
                                "The compressed-RTF stream ends inside a literal token.", location);
                            return false;
                        }
                        byte literal = source[inputOffset++];
                        if (rawSize == 0 && !ignoredEmptySentinel && output.Length == 0 && literal == 0) {
                            dictionary.Add(literal);
                            ignoredEmptySentinel = true;
                        } else if (!TryWriteDecodedByte(literal, rawSize, dictionary, output,
                            diagnostics, location)) return false;
                        continue;
                    }

                    if (inputOffset > inputEnd - 2) {
                        AddError(diagnostics, "EMAIL_MSG_RTF_CONTENT_TRUNCATED",
                            "The compressed-RTF stream ends inside a dictionary reference.", location);
                        return false;
                    }
                    int reference = (source[inputOffset] << 8) | source[inputOffset + 1];
                    inputOffset += 2;
                    int readOffset = reference >> 4;
                    if (readOffset == dictionary.WriteOffset) {
                        terminated = true;
                        break;
                    }

                    int length = (reference & 0x0f) + 2;
                    for (int index = 0; index < length; index++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        byte value = dictionary.Read(readOffset);
                        readOffset = (readOffset + 1) & CompressionDictionary.OffsetMask;
                        if (!TryWriteDecodedByte(value, rawSize, dictionary, output,
                            diagnostics, location)) return false;
                    }
                }
            }

            if (!terminated) {
                AddError(diagnostics, "EMAIL_MSG_RTF_TERMINATOR_MISSING",
                    "The compressed-RTF stream has no completion dictionary reference.", location);
                return false;
            }
            if (output.Length != rawSize) {
                AddError(diagnostics, "EMAIL_MSG_RTF_SIZE_MISMATCH",
                    string.Concat("The compressed RTF produced ", output.Length.ToString(CultureInfo.InvariantCulture),
                        " bytes instead of the declared ", rawSize.ToString(CultureInfo.InvariantCulture), "."), location);
                return false;
            }
            result = output.ToArray();
            return true;
        }
    }

    internal static byte[] Compress(byte[] source, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        var dictionary = new CompressionDictionary();
        using (var contents = new MemoryStream()) {
            int inputOffset = 0;
            bool completed = false;
            while (!completed) {
                cancellationToken.ThrowIfCancellationRequested();
                byte control = 0;
                byte[] tokens = new byte[16];
                int tokenOffset = 0;
                for (int token = 0; token < 8; token++) {
                    int controlBit = 1 << token;
                    if (inputOffset >= source.Length) {
                        if (source.Length == 0 && token == 0) {
                            tokens[tokenOffset++] = 0;
                            dictionary.Add(0);
                            continue;
                        }
                        control = unchecked((byte)(control | controlBit));
                        WriteReference(tokens, ref tokenOffset, dictionary.WriteOffset, 0);
                        completed = true;
                        break;
                    }

                    Match match = dictionary.FindLongestMatch(source, inputOffset);
                    if (match.Length >= 2) {
                        control = unchecked((byte)(control | controlBit));
                        WriteReference(tokens, ref tokenOffset, match.Offset, match.Length);
                        for (int index = 0; index < match.Length; index++) dictionary.Add(source[inputOffset + index]);
                        inputOffset += match.Length;
                    } else {
                        byte value = source[inputOffset++];
                        tokens[tokenOffset++] = value;
                        dictionary.Add(value);
                    }
                }
                contents.WriteByte(control);
                contents.Write(tokens, 0, tokenOffset);
            }

            byte[] contentBytes = contents.ToArray();
            byte[] result = new byte[checked(HeaderLength + contentBytes.Length)];
            MsgBinary.WriteUInt32(result, 0, checked((uint)(contentBytes.Length + 12)));
            MsgBinary.WriteUInt32(result, 4, checked((uint)source.Length));
            MsgBinary.WriteUInt32(result, 8, CompressedType);
            MsgBinary.WriteUInt32(result, 12, CalculateCrc(contentBytes, 0, contentBytes.Length));
            Buffer.BlockCopy(contentBytes, 0, result, HeaderLength, contentBytes.Length);
            return result;
        }
    }

    internal static uint CalculateCrc(byte[] bytes, int offset, int count) {
        uint crc = 0;
        int end = checked(offset + count);
        for (int index = offset; index < end; index++) {
            crc ^= bytes[index];
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320U : crc >> 1;
            }
        }
        return crc;
    }

    private static bool TryWriteDecodedByte(byte value, uint rawSize, CompressionDictionary dictionary,
        MemoryStream output, IList<EmailDiagnostic> diagnostics, string location) {
        if (output.Length >= rawSize) {
            AddError(diagnostics, "EMAIL_MSG_RTF_SIZE_MISMATCH",
                "The compressed RTF expands beyond its declared raw size.", location);
            return false;
        }
        output.WriteByte(value);
        dictionary.Add(value);
        return true;
    }

    private static void WriteReference(byte[] tokens, ref int tokenOffset, int offset, int length) {
        int encodedLength = length == 0 ? 0 : length - 2;
        int reference = (offset << 4) | encodedLength;
        tokens[tokenOffset++] = unchecked((byte)(reference >> 8));
        tokens[tokenOffset++] = unchecked((byte)reference);
    }

    private static void AddError(IList<EmailDiagnostic> diagnostics, string code, string message, string location) {
        diagnostics.Add(new EmailDiagnostic(code, message, EmailDiagnosticSeverity.Error, location));
    }

    private readonly struct Match {
        internal Match(int offset, int length) {
            Offset = offset;
            Length = length;
        }

        internal int Offset { get; }
        internal int Length { get; }
    }

    private sealed class CompressionDictionary {
        internal const int OffsetMask = 4095;
        private const int BufferLength = 4096;
        private const int MaximumMatchLength = 17;
        private const string InitialContents =
            "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx";

        private readonly byte[] _buffer = new byte[BufferLength];
        private readonly bool[] _valid = new bool[BufferLength];
        private readonly int[] _next = new int[BufferLength];
        private readonly int[] _previous = new int[BufferLength];
        private readonly int[] _heads = new int[256];
        private readonly int[] _tails = new int[256];

        internal CompressionDictionary() {
            FillWithMinusOne(_next);
            FillWithMinusOne(_previous);
            FillWithMinusOne(_heads);
            FillWithMinusOne(_tails);
            byte[] initial = Encoding.ASCII.GetBytes(InitialContents);
            if (initial.Length != 207) throw new InvalidOperationException("The compressed-RTF dictionary seed is invalid.");
            for (int index = 0; index < initial.Length; index++) {
                _buffer[index] = initial[index];
                _valid[index] = true;
                Insert(index, initial[index]);
            }
            WriteOffset = initial.Length;
            EndOffset = initial.Length;
        }

        internal int WriteOffset { get; private set; }
        internal int EndOffset { get; private set; }

        internal byte Read(int offset) => _buffer[offset & OffsetMask];

        internal void Add(byte value) {
            int position = WriteOffset;
            if (_valid[position]) Remove(position, _buffer[position]);
            _buffer[position] = value;
            _valid[position] = true;
            Insert(position, value);
            if (EndOffset < BufferLength) EndOffset++;
            WriteOffset = (WriteOffset + 1) & OffsetMask;
        }

        internal Match FindLongestMatch(byte[] input, int inputOffset) {
            int maximumLength = Math.Min(MaximumMatchLength, input.Length - inputOffset);
            int bestOffset = 0;
            int bestLength = 0;
            int candidate = _heads[input[inputOffset]];
            while (candidate >= 0) {
                if (candidate != WriteOffset) {
                    int length = MatchLength(candidate, input, inputOffset, maximumLength);
                    if (length > bestLength) {
                        bestOffset = candidate;
                        bestLength = length;
                        if (bestLength == maximumLength) break;
                    }
                }
                candidate = _next[candidate];
            }
            return new Match(bestOffset, bestLength);
        }

        private int MatchLength(int candidate, byte[] input, int inputOffset, int maximumLength) {
            int length = 0;
            while (length < maximumLength) {
                int sourceOffset = (candidate + length) & OffsetMask;
                if (!TryReadForMatch(sourceOffset, length, input, inputOffset, out byte value) ||
                    value != input[inputOffset + length]) break;
                length++;
            }
            return length;
        }

        private bool TryReadForMatch(int sourceOffset, int matchedLength, byte[] input, int inputOffset,
            out byte value) {
            for (int written = matchedLength - 1; written >= 0; written--) {
                if (((WriteOffset + written) & OffsetMask) == sourceOffset) {
                    value = input[inputOffset + written];
                    return true;
                }
            }
            value = _buffer[sourceOffset];
            return _valid[sourceOffset];
        }

        private void Insert(int position, byte value) {
            int tail = _tails[value];
            _previous[position] = tail;
            _next[position] = -1;
            if (tail < 0) _heads[value] = position;
            else _next[tail] = position;
            _tails[value] = position;
        }

        private void Remove(int position, byte value) {
            int previous = _previous[position];
            int next = _next[position];
            if (previous < 0) _heads[value] = next;
            else _next[previous] = next;
            if (next < 0) _tails[value] = previous;
            else _previous[next] = previous;
            _previous[position] = -1;
            _next[position] = -1;
        }

        private static void FillWithMinusOne(int[] values) {
            for (int index = 0; index < values.Length; index++) values[index] = -1;
        }
    }
}

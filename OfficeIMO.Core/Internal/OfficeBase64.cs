#if NET8_0_OR_GREATER
using System.Buffers;
using System.Buffers.Text;
#endif

namespace OfficeIMO.Core.Internal;

/// <summary>Shared byte-oriented base64 kernels and MIME whitespace/padding handling.</summary>
internal static class OfficeBase64 {
    private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
#if NET8_0_OR_GREATER
    private static readonly SearchValues<byte> AlphabetBytes = SearchValues.Create(
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"u8);
    private static readonly SearchValues<byte> MimeWhiteSpace = SearchValues.Create(" \t\r\n\f\v"u8);
#endif

    /// <summary>Decodes a base64 string suffix with Convert's strict padding and four-character whitespace contract.</summary>
    internal static byte[] FromBase64String(string value, int offset) {
#if NET8_0_OR_GREATER
        ReadOnlySpan<char> input = value.AsSpan(offset);
        if (Base64.IsValid(input, out int length)) {
            if (length == 0) return Array.Empty<byte>();
            var output = new byte[length];
            if (Convert.TryFromBase64Chars(input, output, out int written) && written == length) return output;
        }
#endif
        // Keep Convert's exception details on malformed data and its scalar implementation on older targets.
        return Convert.FromBase64String(value.Substring(offset));
    }

    /// <summary>Incrementally validates MIME base64 without allocating or decoding payloads.</summary>
    internal struct MimeAnalysis {
        private long _characters;
        private int _padding;
        private bool _invalid;

        internal void Append(byte[] input, int offset, int count) {
            if (_invalid) return;
            int end = checked(offset + count);
            while (offset < end) {
#if NET8_0_OR_GREATER
                if (_padding == 0) {
                    int run = input.AsSpan(offset, end - offset).IndexOfAnyExcept(AlphabetBytes);
                    if (run < 0) run = end - offset;
                    _characters += run;
                    offset += run;
                    if (offset == end) break;
                }
#endif
                byte value = input[offset++];
                if (IsMimeWhiteSpace(value)) {
                    while (offset < end && IsMimeWhiteSpace(input[offset])) offset++;
                    continue;
                }
                _characters++;
                if (value == '=') {
                    // Saturating at three keeps arbitrarily long padding invalid across chunks.
                    if (_padding < 3) _padding++;
                } else if (_padding != 0 || DecodeValue(value) < 0) {
                    _invalid = true;
                    return;
                }
            }
        }

        internal bool HasInvalidData => _invalid;
        internal bool IsValid => !_invalid && _padding <= 2 &&
            (_padding == 0 ? _characters % 4 != 1 : _characters % 4 == 0);
        internal bool RecoveredPadding => IsValid && _padding == 0 && _characters % 4 != 0;
        internal long DecodedLength => _characters / 4 * 3 - _padding +
            (_padding == 0 ? _characters % 4 == 2 ? 1 : _characters % 4 == 3 ? 2 : 0 : 0);
    }

    /// <summary>Returns the decoded MIME length, or the original length when malformed data must be preserved.</summary>
    internal static long GetMimeDecodedLength(byte[] input, int offset, int count,
        out bool valid, out bool recoveredPadding) {
        var analysis = new MimeAnalysis();
        analysis.Append(input, offset, count);
        valid = analysis.IsValid;
        recoveredPadding = analysis.RecoveredPadding;
        return valid ? analysis.DecodedLength : count;
    }

    /// <summary>Copies MIME base64 without its six ASCII whitespace characters. In-place compaction is supported.</summary>
    internal static int CompactMime(byte[] input, int offset, int count, byte[] output, int outputOffset) {
        int start = outputOffset;
        int end = checked(offset + count);
        while (offset < end) {
            int run;
#if NET8_0_OR_GREATER
            run = input.AsSpan(offset, end - offset).IndexOfAny(MimeWhiteSpace);
            if (run < 0) run = end - offset;
#else
            int next = offset;
            while (next < end && !IsMimeWhiteSpace(input[next])) next++;
            run = next - offset;
#endif
            if (run != 0) {
                Buffer.BlockCopy(input, offset, output, outputOffset, run);
                outputOffset += run;
                offset += run;
            }
            while (offset < end && IsMimeWhiteSpace(input[offset])) offset++;
        }
        return outputOffset - start;
    }

    /// <summary>Decodes MIME bytes into bounded caller-owned storage, recovering missing terminal padding.</summary>
    internal static bool TryDecodeMime(byte[] input, int offset, int count, byte[] output,
        out int written, out bool recoveredPadding) {
        long length = GetMimeDecodedLength(input, offset, count, out bool valid, out recoveredPadding);
        written = 0;
        if (!valid || length > output.Length) return false;
#if NET8_0_OR_GREATER
        byte[] compact = ArrayPool<byte>.Shared.Rent(checked(count + 3));
        try {
            int compactLength = CompactMime(input, offset, count, compact, 0);
            int padding = (4 - compactLength % 4) % 4;
            recoveredPadding = padding != 0;
            for (int index = 0; index < padding; index++) compact[compactLength++] = (byte)'=';
            return TryDecodeBlock(compact, compactLength, output, out written);
        } finally {
            ArrayPool<byte>.Shared.Return(compact);
        }
#else
        return DecodeScalar(input, offset, count, output, true, out written, out recoveredPadding);
#endif
    }

    /// <summary>Decodes a whitespace-free padded block, as produced by MIME compaction or an ordinary encoder.</summary>
    internal static bool TryDecodeBlock(byte[] input, int count, byte[] output, out int written) {
#if NET8_0_OR_GREATER
        OperationStatus status = Base64.DecodeFromUtf8(input.AsSpan(0, count), output,
            out int consumed, out written);
        return status == OperationStatus.Done && consumed == count;
#else
        return DecodeScalar(input, 0, count, output, false, out written, out _);
#endif
    }

    /// <summary>Encodes one complete byte range to ASCII without an intermediate string or character array.</summary>
    internal static int EncodeToUtf8(byte[] input, int offset, int count, byte[] output, int outputOffset) {
#if NET8_0_OR_GREATER
        OperationStatus status = Base64.EncodeToUtf8(input.AsSpan(offset, count), output.AsSpan(outputOffset),
            out int consumed, out int written);
        if (status != OperationStatus.Done || consumed != count)
            throw new ArgumentException("The Base64 output buffer is too small.", nameof(output));
        return written;
#else
        int required = checked((count + 2) / 3 * 4);
        if (outputOffset < 0 || outputOffset > output.Length - required)
            throw new ArgumentException("The Base64 output buffer is too small.", nameof(output));
        int start = outputOffset;
        int end = checked(offset + count);
        while (offset + 2 < end) {
            int value = input[offset++] << 16 | input[offset++] << 8 | input[offset++];
            output[outputOffset++] = (byte)Alphabet[value >> 18];
            output[outputOffset++] = (byte)Alphabet[value >> 12 & 63];
            output[outputOffset++] = (byte)Alphabet[value >> 6 & 63];
            output[outputOffset++] = (byte)Alphabet[value & 63];
        }
        if (offset < end) {
            int first = input[offset++];
            int second = offset < end ? input[offset] : -1;
            output[outputOffset++] = (byte)Alphabet[first >> 2];
            output[outputOffset++] = (byte)Alphabet[(first & 3) << 4 | (second < 0 ? 0 : second >> 4)];
            output[outputOffset++] = second < 0 ? (byte)'=' : (byte)Alphabet[(second & 15) << 2];
            output[outputOffset++] = (byte)'=';
        }
        return outputOffset - start;
#endif
    }

    private static bool DecodeScalar(byte[] input, int offset, int count, byte[] output, bool mime,
        out int written, out bool recoveredPadding) {
        written = 0;
        recoveredPadding = false;
        int group = 0, first = 0, second = 0, third = 0;
        bool sawPadding = false;
        int end = checked(offset + count);
        for (int index = offset; index < end; index++) {
            byte current = input[index];
            if (mime && IsMimeWhiteSpace(current)) continue;
            int value;
            if (current == '=') {
                value = -2;
                sawPadding = true;
            } else {
                value = DecodeValue(current);
                if (sawPadding || value < 0) return false;
            }
            switch (group++) {
                case 0: first = value; break;
                case 1: second = value; break;
                case 2: third = value; break;
                default:
                    if (!WriteQuartet(output, ref written, first, second, third, value)) return false;
                    group = 0;
                    break;
            }
        }
        if (group == 0) return true;
        if (!mime || sawPadding || group == 1) return false;
        recoveredPadding = true;
        return WriteQuartet(output, ref written, first, second, group == 2 ? -2 : third, -2);
    }

    private static bool WriteQuartet(byte[] output, ref int written, int first, int second, int third, int fourth) {
        if (first < 0 || second < 0 || third == -2 && fourth != -2 || written >= output.Length) return false;
        output[written++] = (byte)(first << 2 | second >> 4);
        if (third == -2) return true;
        if (third < 0 || written >= output.Length) return false;
        output[written++] = (byte)(second << 4 | third >> 2);
        if (fourth == -2) return true;
        if (fourth < 0 || written >= output.Length) return false;
        output[written++] = (byte)(third << 6 | fourth);
        return true;
    }

    private static int DecodeValue(byte value) {
        if (value >= 'A' && value <= 'Z') return value - 'A';
        if (value >= 'a' && value <= 'z') return value - 'a' + 26;
        if (value >= '0' && value <= '9') return value - '0' + 52;
        if (value == '+') return 62;
        if (value == '/') return 63;
        return -1;
    }

    private static bool IsMimeWhiteSpace(byte value) =>
        value == ' ' || value == '\t' || value == '\r' || value == '\n' || value == '\f' || value == '\v';
}

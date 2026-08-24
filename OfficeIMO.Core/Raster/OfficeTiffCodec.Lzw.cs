using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    private const int LzwClearCode = 256;
    private const int LzwEndCode = 257;
    private const int LzwFirstCode = 258;
    private const int LzwMaximumCode = 4095;

    private static byte[] EncodeLzw(byte[] input, int inputCount) {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (inputCount < 0 || inputCount > input.Length) throw new ArgumentOutOfRangeException(nameof(inputCount));
        using var writer = new TiffLzwBitWriter(Math.Max(32, inputCount));
        writer.Write(LzwClearCode, 9);
        if (inputCount == 0) {
            writer.Write(LzwEndCode, 9);
            return writer.Finish();
        }

        var dictionary = new Dictionary<int, int>(Math.Min(4096, inputCount));
        int codeSize = 9;
        int nextCode = LzwFirstCode;
        int prefix = input[0];
        for (int index = 1; index < inputCount; index++) {
            byte suffix = input[index];
            int key = (prefix << 8) | suffix;
            if (dictionary.TryGetValue(key, out int combined)) {
                prefix = combined;
                continue;
            }

            writer.Write(prefix, codeSize);
            if (nextCode <= LzwMaximumCode) {
                dictionary.Add(key, nextCode++);
                if (codeSize < 12 && nextCode == 1 << codeSize) codeSize++;
            } else {
                writer.Write(LzwClearCode, codeSize);
                dictionary.Clear();
                codeSize = 9;
                nextCode = LzwFirstCode;
            }
            prefix = suffix;
        }
        writer.Write(prefix, codeSize);
        writer.Write(LzwEndCode, codeSize);
        return writer.Finish();
    }

    private static bool TryDecodeLzw(
        byte[] input,
        int inputOffset,
        int inputCount,
        byte[] output,
        int outputOffset,
        int expectedCount,
        CancellationToken cancellationToken) {
        if (expectedCount <= 0 || outputOffset < 0 || outputOffset > output.Length - expectedCount) return false;
        var reader = new TiffLzwBitReader(input, inputOffset, inputCount);
        var prefixes = new short[4096];
        var suffixes = new byte[4096];
        var stack = new byte[4096];
        int codeSize = 9;
        int nextCode = LzwFirstCode;
        int previousCode = -1;
        byte previousFirst = 0;
        int target = outputOffset;
        int outputEnd = checked(outputOffset + expectedCount);
        bool ended = false;
        bool requiresInitialClear = true;
        int decodedCodes = 0;

        while (reader.TryRead(codeSize, out int code)) {
            if ((decodedCodes++ & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (requiresInitialClear) {
                requiresInitialClear = false;
                if (code != LzwClearCode) return false;
            }
            if (code == LzwClearCode) {
                codeSize = 9;
                nextCode = LzwFirstCode;
                previousCode = -1;
                continue;
            }
            if (code == LzwEndCode) {
                ended = true;
                break;
            }
            if (code < 0 || code > LzwMaximumCode || code > nextCode || previousCode < 0 && code >= 256) return false;

            int expandedCode = code;
            int stackCount = 0;
            if (code == nextCode) {
                if (previousCode < 0) return false;
                stack[stackCount++] = previousFirst;
                expandedCode = previousCode;
            }
            while (expandedCode >= 256) {
                if (expandedCode < LzwFirstCode || expandedCode >= nextCode || stackCount >= stack.Length) return false;
                stack[stackCount++] = suffixes[expandedCode];
                expandedCode = prefixes[expandedCode];
            }
            byte first = (byte)expandedCode;
            if (stackCount >= stack.Length) return false;
            stack[stackCount++] = first;
            if (target > outputEnd - stackCount) return false;
            for (int index = stackCount - 1; index >= 0; index--) output[target++] = stack[index];

            if (previousCode >= 0 && nextCode <= LzwMaximumCode) {
                prefixes[nextCode] = (short)previousCode;
                suffixes[nextCode] = first;
                nextCode++;
                // The decoder constructs each dictionary entry one emitted code later than the encoder.
                if (codeSize < 12 && nextCode == (1 << codeSize) - 1) codeSize++;
            }
            previousCode = code;
            previousFirst = first;
        }
        return ended && target == outputEnd && reader.HasOnlyZeroBitPaddingAtEnd;
    }

    private sealed class TiffLzwBitWriter : IDisposable {
        private byte[] _buffer;
        private int _length;
        private uint _bits;
        private int _bitCount;

        internal TiffLzwBitWriter(int capacity) {
#if NET8_0_OR_GREATER
            _buffer = ArrayPool<byte>.Shared.Rent(capacity);
#else
            _buffer = new byte[capacity];
#endif
        }

        internal void Write(int code, int width) {
            _bits = (_bits << width) | (uint)code;
            _bitCount += width;
            while (_bitCount >= 8) {
                EnsureCapacity();
                _bitCount -= 8;
                _buffer[_length++] = (byte)(_bits >> _bitCount);
                _bits &= _bitCount == 0 ? 0U : (1U << _bitCount) - 1U;
            }
        }

        internal byte[] Finish() {
            if (_bitCount > 0) {
                EnsureCapacity();
                _buffer[_length++] = (byte)(_bits << (8 - _bitCount));
            }
            var result = new byte[_length];
            Buffer.BlockCopy(_buffer, 0, result, 0, _length);
            return result;
        }

        public void Dispose() {
#if NET8_0_OR_GREATER
            byte[] buffer = _buffer;
            _buffer = Array.Empty<byte>();
            if (buffer.Length > 0) ArrayPool<byte>.Shared.Return(buffer);
#endif
        }

        private void EnsureCapacity() {
            if (_length < _buffer.Length) return;
#if NET8_0_OR_GREATER
            byte[] expanded = ArrayPool<byte>.Shared.Rent(checked(_buffer.Length * 2));
            Buffer.BlockCopy(_buffer, 0, expanded, 0, _length);
            ArrayPool<byte>.Shared.Return(_buffer);
            _buffer = expanded;
#else
            Array.Resize(ref _buffer, checked(_buffer.Length * 2));
#endif
        }
    }

    private sealed class TiffLzwBitReader {
        private readonly byte[] _input;
        private readonly int _end;
        private int _offset;
        private uint _bits;
        private int _bitCount;

        internal TiffLzwBitReader(byte[] input, int offset, int count) {
            _input = input;
            _offset = offset;
            _end = checked(offset + count);
        }

        internal bool TryRead(int width, out int value) {
            while (_bitCount < width) {
                if (_offset >= _end) {
                    value = 0;
                    return false;
                }
                _bits = (_bits << 8) | _input[_offset++];
                _bitCount += 8;
            }
            _bitCount -= width;
            value = (int)((_bits >> _bitCount) & ((1U << width) - 1U));
            _bits &= _bitCount == 0 ? 0U : (1U << _bitCount) - 1U;
            return true;
        }

        internal bool HasOnlyZeroBitPaddingAtEnd => _offset == _end && _bitCount < 8 && _bits == 0U;
    }
}

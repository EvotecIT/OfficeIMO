using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    private const int Vp8lCompressionMaximumPixels = 4_000_000;
    private const int Vp8lMatchTableMaximumEntries = 65_536;

    private static byte[]? TryEncodeCompressedVp8l(
        int width,
        int height,
        bool hasAlpha,
        byte[] rgba) {
        try {
            // The literal encoder remains available for larger images. Keep the optional
            // compression candidate below a separate allocation ceiling because it owns
            // residual pixels and a dynamic bit stream at the same time.
            if (rgba.Length / 4 > Vp8lCompressionMaximumPixels) return null;
            var residuals = CreateVp8lResiduals(width, height, rgba);
            var frequencies = new Vp8lEncodingFrequencies();
            var lastPosition = new Dictionary<uint, int>();
            VisitVp8lEncodingTokens(residuals, lastPosition, (color, length, distance) => {
                if (distance == 0) {
                    frequencies.Green[(int)(color >> 8) & 255]++;
                    frequencies.Red[(int)(color >> 16) & 255]++;
                    frequencies.Blue[(int)color & 255]++;
                    frequencies.Alpha[(int)(color >> 24) & 255]++;
                } else {
                    GetVp8lPrefix(length, 24, out int prefix, out _, out _);
                    GetVp8lPrefix(distance + 120, 40, out int distancePrefix, out _, out _);
                    frequencies.Green[256 + prefix]++;
                    frequencies.Distance[distancePrefix]++;
                }
            });
            var writer = new DynamicLsbBitWriter(Math.Max(128, rgba.Length / 4));
            writer.WriteBits((uint)(width - 1), 14);
            writer.WriteBits((uint)(height - 1), 14);
            writer.WriteBits(hasAlpha ? 1U : 0U, 1);
            writer.WriteBits(0, 3);

            writer.WriteBits(1, 1);
            writer.WriteBits(0, 2); // predictor transform
            const int predictorSizeBits = 9;
            writer.WriteBits(predictorSizeBits - 2, 3);
            WriteConstantVp8lImage(writer, 12, 0, 0, 255);
            writer.WriteBits(1, 1);
            writer.WriteBits(2, 2); // subtract-green transform
            writer.WriteBits(0, 1); // end transforms
            writer.WriteBits(0, 1); // no color cache
            writer.WriteBits(0, 1); // one prefix group

            byte[] greenLengths = OfficeHuffmanCodeLengths.Create(frequencies.Green, 15);
            byte[] redLengths = OfficeHuffmanCodeLengths.Create(frequencies.Red, 15);
            byte[] blueLengths = OfficeHuffmanCodeLengths.Create(frequencies.Blue, 15);
            byte[] alphaLengths = OfficeHuffmanCodeLengths.Create(frequencies.Alpha, 15);
            byte[] distanceLengths = OfficeHuffmanCodeLengths.Create(frequencies.Distance, 15);
            WriteVp8lHuffmanTree(writer, greenLengths);
            WriteVp8lHuffmanTree(writer, redLengths);
            WriteVp8lHuffmanTree(writer, blueLengths);
            WriteVp8lHuffmanTree(writer, alphaLengths);
            WriteVp8lHuffmanTree(writer, distanceLengths);
            var greenCodes = new Vp8lCodebook(greenLengths);
            var redCodes = new Vp8lCodebook(redLengths);
            var blueCodes = new Vp8lCodebook(blueLengths);
            var alphaCodes = new Vp8lCodebook(alphaLengths);
            var distanceCodes = new Vp8lCodebook(distanceLengths);
            VisitVp8lEncodingTokens(residuals, lastPosition, (color, length, distance) => {
                if (distance == 0) {
                    greenCodes.Write(writer, (int)(color >> 8) & 255);
                    redCodes.Write(writer, (int)(color >> 16) & 255);
                    blueCodes.Write(writer, (int)color & 255);
                    alphaCodes.Write(writer, (int)(color >> 24) & 255);
                } else {
                    GetVp8lPrefix(length, 24, out int prefix, out int extraBits, out int extraValue);
                    greenCodes.Write(writer, 256 + prefix);
                    if (extraBits > 0) writer.WriteBits((uint)extraValue, extraBits);
                    GetVp8lPrefix(distance + 120, 40, out int distancePrefix, out int distanceExtraBits, out int distanceExtraValue);
                    distanceCodes.Write(writer, distancePrefix);
                    if (distanceExtraBits > 0) writer.WriteBits((uint)distanceExtraValue, distanceExtraBits);
                }
            });
            byte[] bits = writer.Finish();
            var payload = new byte[bits.Length + 1];
            payload[0] = 0x2F;
            Buffer.BlockCopy(bits, 0, payload, 1, bits.Length);
            return payload;
        } catch (OverflowException) {
            return null;
        }
    }

    // Replaying the deterministic match walk avoids a pixel-sized token buffer.
    // The same bounded dictionary is cleared and reused between frequency and emission passes.
    private static void VisitVp8lEncodingTokens(uint[] residuals, Dictionary<uint, int> lastPosition,
        Action<uint, int, int> visit) {
        lastPosition.Clear();
        int position = 0;
        while (position < residuals.Length) {
            int matchLength = 0;
            int matchDistance = 0;
            uint current = residuals[position];
            if (lastPosition.TryGetValue(current, out int previous)) {
                int distance = position - previous;
                int maximum = Math.Min(4096, residuals.Length - position);
                while (matchLength < maximum &&
                       residuals[position + matchLength] == residuals[position + matchLength - distance]) {
                    matchLength++;
                }
                if (matchLength >= 3 && CanEncodeVp8lPrefix(checked(distance + 120), 40)) {
                    matchDistance = distance;
                }
            }
            if (matchDistance > 0) {
                visit(0, matchLength, matchDistance);
                for (int index = 0; index < matchLength; index++) {
                    RememberVp8lPosition(lastPosition, residuals[position + index], position + index);
                }
                position += matchLength;
                continue;
            }

            uint color = residuals[position];
            visit(color, 1, 0);
            RememberVp8lPosition(lastPosition, color, position++);
        }
    }

    private sealed class Vp8lEncodingFrequencies {
        internal readonly int[] Green = new int[280];
        internal readonly int[] Red = new int[256];
        internal readonly int[] Blue = new int[256];
        internal readonly int[] Alpha = new int[256];
        internal readonly int[] Distance = new int[40];
    }

    private static void RememberVp8lPosition(Dictionary<uint, int> lastPosition, uint color, int position) {
        if (lastPosition.Count >= Vp8lMatchTableMaximumEntries && !lastPosition.ContainsKey(color)) {
            lastPosition.Clear();
        }
        lastPosition[color] = position;
    }

    private static uint[] CreateVp8lResiduals(int width, int height, byte[] rgba) {
        var predicted = new uint[checked(width * height)];
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                int position = y * width + x;
                int offset = position * 4;
                uint color = (uint)(rgba[offset + 3] << 24 | rgba[offset] << 16 | rgba[offset + 1] << 8 | rgba[offset + 2]);
                uint predictor = x == 0 && y == 0
                    ? 0xFF000000U
                    : y == 0 ? PackRgba(rgba, offset - 4)
                    : x == 0 ? PackRgba(rgba, offset - width * 4)
                    : PredictVp8l(12,
                        PackRgba(rgba, offset - 4), PackRgba(rgba, offset - width * 4),
                        PackRgba(rgba, offset - width * 4 - 4), 0U); // mode 12 does not use top-right
                uint residual = SubtractArgb(color, predictor);
                int green = (int)(residual >> 8) & 255;
                int red = (((int)(residual >> 16) & 255) - green) & 255;
                int blue = (((int)residual & 255) - green) & 255;
                predicted[position] = (residual & 0xFF00FF00U) | (uint)(red << 16 | blue);
            }
        }
        return predicted;
    }

    private static uint PackRgba(byte[] rgba, int offset) =>
        (uint)(rgba[offset + 3] << 24 | rgba[offset] << 16 | rgba[offset + 1] << 8 | rgba[offset + 2]);

    private static uint SubtractArgb(uint value, uint predictor) {
        uint result = 0;
        for (int shift = 0; shift <= 24; shift += 8) {
            result |= (uint)((((int)(value >> shift) - (int)(predictor >> shift)) & 255)) << shift;
        }
        return result;
    }

    private static void WriteConstantVp8lImage(
        ILsbBitWriter writer,
        int green,
        int red,
        int blue,
        int alpha) {
        writer.WriteBits(0, 1); // no cache; subimages do not carry the meta-code flag
        WriteSingleSymbolTree(writer, green);
        WriteSingleSymbolTree(writer, red);
        WriteSingleSymbolTree(writer, blue);
        WriteSingleSymbolTree(writer, alpha);
        WriteSingleSymbolTree(writer, 0);
    }

    private static void WriteSingleSymbolTree(ILsbBitWriter writer, int symbol) {
        writer.WriteBits(1, 1);
        writer.WriteBits(0, 1);
        writer.WriteBits(symbol > 1 ? 1U : 0U, 1);
        writer.WriteBits((uint)symbol, symbol > 1 ? 8 : 1);
    }

    private static void WriteVp8lHuffmanTree(ILsbBitWriter writer, byte[] lengths) {
        int symbols = 0, singleSymbol = -1;
        for (int symbol = 0; symbol < lengths.Length; symbol++) if (lengths[symbol] > 0) { symbols++; singleSymbol = symbol; }
        if (symbols == 1 && singleSymbol <= 255) {
            WriteSingleSymbolTree(writer, singleSymbol);
            return;
        }
        byte[] codeLengthLengths = CreateMixedDepthLengths(19, 13, 4, 5);
        var codeLengthCodes = new Vp8lCodebook(codeLengthLengths);
        writer.WriteBits(0, 1);
        writer.WriteBits(15, 4);
        for (int index = 0; index < Vp8lCodeLengthOrder.Length; index++) {
            writer.WriteBits(codeLengthLengths[Vp8lCodeLengthOrder[index]], 3);
        }
        writer.WriteBits(0, 1);
        for (int index = 0; index < lengths.Length; index++) codeLengthCodes.Write(writer, lengths[index]);
    }

    private static byte[] CreateMixedDepthLengths(int count, int shorterCount, int shortDepth, int longDepth) {
        var lengths = new byte[count];
        for (int index = 0; index < count; index++) lengths[index] = (byte)(index < shorterCount ? shortDepth : longDepth);
        return lengths;
    }

    private static void GetVp8lPrefix(
        int value,
        int prefixCount,
        out int prefix,
        out int extraBits,
        out int extraValue) {
        if (value <= 4) {
            prefix = value - 1;
            extraBits = 0;
            extraValue = 0;
            return;
        }
        for (prefix = 4; prefix < prefixCount; prefix++) {
            extraBits = (prefix - 2) >> 1;
            int first = ((2 + (prefix & 1)) << extraBits) + 1;
            int last = first + (1 << extraBits) - 1;
            if (value <= last) {
                extraValue = value - first;
                return;
            }
        }
        throw new ArgumentOutOfRangeException(nameof(value));
    }

    private static bool CanEncodeVp8lPrefix(int value, int prefixCount) {
        if (value <= 4) return value >= 1;
        int prefix = prefixCount - 1;
        int extraBits = (prefix - 2) >> 1;
        int first = ((2 + (prefix & 1)) << extraBits) + 1;
        int last = first + (1 << extraBits) - 1;
        return value <= last;
    }

    private sealed class Vp8lCodebook {
        private readonly uint[] _codes;
        private readonly byte[] _lengths;
        private readonly int _singleSymbol;

        internal Vp8lCodebook(byte[] lengths) {
            _lengths = lengths;
            _codes = new uint[lengths.Length];
            var counts = new int[16];
            for (int index = 0; index < lengths.Length; index++) counts[lengths[index]]++;
            int symbols = lengths.Length - counts[0];
            counts[0] = 0;
            _singleSymbol = -1;
            var next = new int[16];
            int code = 0;
            for (int length = 1; length <= 15; length++) {
                code = (code + counts[length - 1]) << 1;
                next[length] = code;
            }
            for (int symbol = 0; symbol < lengths.Length; symbol++) {
                int length = lengths[symbol];
                if (length == 0) continue;
                if (symbols == 1) _singleSymbol = symbol;
                _codes[symbol] = ReverseBits((uint)next[length]++, length);
            }
        }

        internal void Write(ILsbBitWriter writer, int symbol) {
            if (symbol == _singleSymbol) return;
            int length = _lengths[symbol];
            if (length == 0) throw new InvalidOperationException("Cannot emit an unused WebP Huffman symbol.");
            writer.WriteBits(_codes[symbol], length);
        }
    }

    private static uint ReverseBits(uint value, int length) {
        uint reversed = 0;
        for (int index = 0; index < length; index++) {
            reversed = (reversed << 1) | (value & 1U);
            value >>= 1;
        }
        return reversed;
    }

    private sealed class DynamicLsbBitWriter : ILsbBitWriter {
        private readonly List<byte> _bytes;
        private ulong _buffer;
        private int _bitCount;

        internal DynamicLsbBitWriter(int capacity) => _bytes = new List<byte>(capacity);

        public void WriteBits(uint value, int count) {
            ulong mask = count == 32 ? uint.MaxValue : (1UL << count) - 1UL;
            _buffer |= ((ulong)value & mask) << _bitCount;
            _bitCount += count;
            while (_bitCount >= 8) {
                _bytes.Add((byte)_buffer);
                _buffer >>= 8;
                _bitCount -= 8;
            }
        }

        public void Flush() {
            if (_bitCount == 0) return;
            _bytes.Add((byte)_buffer);
            _buffer = 0;
            _bitCount = 0;
        }

        internal byte[] Finish() {
            Flush();
            return _bytes.ToArray();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    private const int Vp8lCompressionMaximumPixels = 4_000_000;
    private const int Vp8lMatchTableMaximumEntries = 65_536;

    private static byte[]? TryEncodeCompressedVp8l(
        int width,
        int height,
        bool hasAlpha,
        byte[] rgba,
        int literalFileLength,
        long metadataBytes,
        long additionalRetainedManagedBytes,
        CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null,
        OfficeImageEncodingEffort effort = OfficeImageEncodingEffort.Balanced) {
        try {
            // Compression is optional. Keep one pixel workspace, bounded match chains and
            // transform maps; reject the candidate before allocating if its growth peak
            // cannot coexist with the source, metadata and caller-retained buffers.
            if (rgba.Length / 4 > Vp8lCompressionMaximumPixels) return null;
            cancellationToken.ThrowIfCancellationRequested();
            int blockCount = checked(DivideRoundUp(width, 32) * DivideRoundUp(height, 32));
            int maximumBitsBytes = checked((int)((rgba.LongLength * 15L + blockCount * 120L + 7L) / 8L) + 8192);
            int maximumCapacity = Math.Max(128, rgba.Length / 4);
            while (maximumCapacity < maximumBitsBytes) maximumCapacity = checked(maximumCapacity * 2);
            long compressionBytes = checked(rgba.LongLength + 24L +
                Vp8lMatchTableMaximumEntries * 8L + blockCount * 8L + 256L * 1024L +
                maximumCapacity * 3L + 2L * (maximumBitsBytes + 24L));
            if (!IsEncodingWorkingSetWithinLimit(rgba.LongLength, literalFileLength, compressionBytes,
                    metadataBytes, additionalRetainedManagedBytes)) return null;

            var residuals = new uint[rgba.Length / 4];
            var matches = new Vp8lEncodingMatches(effort == OfficeImageEncodingEffort.Fast ? 1 : effort == OfficeImageEncodingEffort.Balanced ? 4 : 16);
            FillVp8lSourcePixels(rgba, residuals, cancellationToken, checkpointObserver);
            bool usePrediction = false, subtractGreen;
            int mapWidth = DivideRoundUp(width, 32);
            uint[] predictors = Array.Empty<uint>(), colors = Array.Empty<uint>();
            Vp8lEncodingPlan plan;
            Vp8lEncodingPlan? predictorPlan = null, colorPlan = null;
            if (effort != OfficeImageEncodingEffort.Smallest) {
                usePrediction = effort == OfficeImageEncodingEffort.Balanced && ShouldUseVp8lPrediction(width, rgba, cancellationToken);
                if (usePrediction) {
                    predictors = CreateVp8lBlockResiduals(width, height, rgba, residuals, cancellationToken, quick: true);
                    predictorPlan = CreateVp8lEncodingPlan(predictors, mapWidth, matches, 0, cancellationToken);
                }
                subtractGreen = effort == OfficeImageEncodingEffort.Balanced && SelectVp8lSubtractGreen(residuals, cancellationToken);
                plan = CreateVp8lEncodingPlan(residuals, width, matches, 8, cancellationToken);
            } else {
                bool rawSubtractGreen = SelectVp8lSubtractGreen(residuals, cancellationToken);
                Vp8lEncodingPlan rawPlan = SelectVp8lEncodingPlan(residuals, width, matches, cancellationToken);
                long rawBits = rawPlan.BitCount + (rawSubtractGreen ? 3 : 0);

                predictors = CreateVp8lBlockResiduals(width, height, rgba, residuals, cancellationToken);
                subtractGreen = SelectVp8lSubtractGreen(residuals, cancellationToken);
                colors = SelectVp8lColorTransform(width, height, residuals, cancellationToken);
                plan = SelectVp8lEncodingPlan(residuals, width, matches, cancellationToken);
                predictorPlan = CreateVp8lEncodingPlan(predictors, mapWidth, matches, 0, cancellationToken);
                colorPlan = colors.Length == 0 ? null
                    : CreateVp8lEncodingPlan(colors, mapWidth, matches, 0, cancellationToken);
                // Subimages omit the main image's meta-prefix flag (one bit).
                long transformedBits = plan.BitCount + 6 + predictorPlan.BitCount - 1 +
                    (subtractGreen ? 3 : 0) + (colorPlan == null ? 0 : 6 + colorPlan.BitCount - 1);
                usePrediction = transformedBits < rawBits;
                if (!usePrediction) {
                    FillVp8lSourcePixels(rgba, residuals, cancellationToken);
                    if (rawSubtractGreen) ApplyVp8lSubtractGreen(residuals, cancellationToken);
                    subtractGreen = rawSubtractGreen;
                    plan = rawPlan;
                }

            }

            var writer = new DynamicLsbBitWriter(Math.Max(128, rgba.Length / 4), maximumBitsBytes);
            writer.WriteBits((uint)(width - 1), 14);
            writer.WriteBits((uint)(height - 1), 14);
            writer.WriteBits(hasAlpha ? 1U : 0U, 1);
            writer.WriteBits(0, 3);
            if (usePrediction) {
                writer.WriteBits(1, 1);
                writer.WriteBits(0, 2); // predictor, applied first during encoding
                writer.WriteBits(3, 3); // 32 x 32 blocks
                WriteVp8lEncodingImage(writer, predictors, mapWidth, matches, predictorPlan!, false, cancellationToken);
            }
            if (subtractGreen) {
                writer.WriteBits(1, 1);
                writer.WriteBits(2, 2);
            }
            if (usePrediction && colorPlan != null) {
                writer.WriteBits(1, 1);
                writer.WriteBits(1, 2); // cross-color, applied to the residual samples
                writer.WriteBits(3, 3);
                WriteVp8lEncodingImage(writer, colors, mapWidth, matches, colorPlan, false, cancellationToken);
            }
            writer.WriteBits(0, 1); // end transforms
            WriteVp8lEncodingImage(writer, residuals, width, matches, plan, true, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] bits = writer.Finish();
            cancellationToken.ThrowIfCancellationRequested();
            var payload = new byte[bits.Length + 1];
            payload[0] = 0x2F;
            for (int offset = 0; offset < bits.Length;) {
                cancellationToken.ThrowIfCancellationRequested();
                int count = Math.Min(16384, bits.Length - offset);
                Buffer.BlockCopy(bits, offset, payload, 1 + offset, count);
                offset += count;
            }
            return payload;
        } catch (OverflowException) {
            return null;
        }
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

    private static void WriteSingleSymbolTree(ILsbBitWriter writer, int symbol) {
        writer.WriteBits(1, 1);
        writer.WriteBits(0, 1);
        writer.WriteBits(symbol > 1 ? 1U : 0U, 1);
        writer.WriteBits((uint)symbol, symbol > 1 ? 8 : 1);
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
        private readonly int _maximumBytes;
        private ulong _buffer;
        private int _bitCount;

        internal DynamicLsbBitWriter(int capacity, int maximumBytes) {
            _bytes = new List<byte>(capacity);
            _maximumBytes = maximumBytes;
        }

        public void WriteBits(uint value, int count) {
            ulong mask = count == 32 ? uint.MaxValue : (1UL << count) - 1UL;
            _buffer |= ((ulong)value & mask) << _bitCount;
            _bitCount += count;
            while (_bitCount >= 8) {
                AddByte((byte)_buffer);
                _buffer >>= 8;
                _bitCount -= 8;
            }
        }

        public void Flush() {
            if (_bitCount == 0) return;
            AddByte((byte)_buffer);
            _buffer = 0;
            _bitCount = 0;
        }

        internal byte[] Finish() {
            Flush();
            return _bytes.ToArray();
        }

        private void AddByte(byte value) {
            if (_bytes.Count >= _maximumBytes) throw new OverflowException("WebP compression exceeded its buffer plan.");
            _bytes.Add(value);
        }
    }
}

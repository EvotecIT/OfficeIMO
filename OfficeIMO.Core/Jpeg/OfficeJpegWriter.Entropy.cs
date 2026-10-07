using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegWriter {
    private static void EncodeBlock(
        BitWriter bw,
        int[] input,
        int[] quant,
        HuffmanTable dcTable,
        HuffmanTable acTable,
        ref int prevDc,
        int[] temp,
        double[] dctScratch) {
        OfficeJpegForwardTransform.Quantize(input, quant, temp, dctScratch);

        var dc = temp[0];
        var diff = dc - prevDc;
        prevDc = dc;
        var dcCat = BitCount(diff);
        bw.WriteBits(dcTable.Codes[dcCat], dcTable.Sizes[dcCat]);
        if (dcCat > 0) {
            bw.WriteBits(EncodeValue(diff, dcCat), dcCat);
        }

        var zeroRun = 0;
        for (var i = 1; i < 64; i++) {
            var v = temp[ZigZag[i]];
            if (v == 0) {
                zeroRun++;
                continue;
            }

            while (zeroRun >= 16) {
                bw.WriteBits(acTable.Codes[0xF0], acTable.Sizes[0xF0]);
                zeroRun -= 16;
            }

            var cat = BitCount(v);
            var symbol = (zeroRun << 4) | cat;
            bw.WriteBits(acTable.Codes[symbol], acTable.Sizes[symbol]);
            bw.WriteBits(EncodeValue(v, cat), cat);
            zeroRun = 0;
        }

        if (zeroRun > 0) {
            bw.WriteBits(acTable.Codes[0x00], acTable.Sizes[0x00]);
        }
    }

    private static void EncodeBlockFromQuantized(
        BitWriter bw,
        short[] coeffs,
        int offset,
        HuffmanTable dcTable,
        HuffmanTable acTable,
        ref int prevDc) {
        var dc = coeffs[offset];
        var diff = dc - prevDc;
        prevDc = dc;
        var dcCat = BitCount(diff);
        bw.WriteBits(dcTable.Codes[dcCat], dcTable.Sizes[dcCat]);
        if (dcCat > 0) {
            bw.WriteBits(EncodeValue(diff, dcCat), dcCat);
        }

        EncodeAcFromQuantized(bw, coeffs, offset, acTable, 1, 63);
    }

    private static void EncodeDcFromQuantized(BitWriter bw, short[] coeffs, int offset, HuffmanTable dcTable, ref int prevDc) {
        var dc = coeffs[offset];
        var diff = dc - prevDc;
        prevDc = dc;
        var dcCat = BitCount(diff);
        bw.WriteBits(dcTable.Codes[dcCat], dcTable.Sizes[dcCat]);
        if (dcCat > 0) {
            bw.WriteBits(EncodeValue(diff, dcCat), dcCat);
        }
    }

    private static void EncodeAcFromQuantized(BitWriter bw, short[] coeffs, int offset, HuffmanTable acTable, int ss, int se) {
        var zeroRun = 0;
        for (var i = ss; i <= se; i++) {
            var v = coeffs[offset + ZigZag[i]];
            if (v == 0) {
                zeroRun++;
                continue;
            }

            while (zeroRun >= 16) {
                bw.WriteBits(acTable.Codes[0xF0], acTable.Sizes[0xF0]);
                zeroRun -= 16;
            }

            var cat = BitCount(v);
            var symbol = (zeroRun << 4) | cat;
            bw.WriteBits(acTable.Codes[symbol], acTable.Sizes[symbol]);
            bw.WriteBits(EncodeValue(v, cat), cat);
            zeroRun = 0;
        }

        if (zeroRun > 0) {
            bw.WriteBits(acTable.Codes[0x00], acTable.Sizes[0x00]);
        }
    }

    private static int BitCount(int value) {
        var v = value < 0 ? -value : value;
        var bits = 0;
        while (v != 0) {
            bits++;
            v >>= 1;
        }
        return bits;
    }

    private static uint EncodeValue(int value, int bits) {
        if (value >= 0) return (uint)value;
        return (uint)(value + (1 << bits) - 1);
    }

    private readonly struct HuffmanTable {
        public readonly ushort[] Codes;
        public readonly byte[] Sizes;
        public HuffmanTable(ushort[] codes, byte[] sizes) {
            Codes = codes;
            Sizes = sizes;
        }
    }

    private sealed class BitWriter {
        private readonly Stream _stream;
        private readonly byte[] _bytes = new byte[4096];
        private int _byteCount;
        private uint _buffer;
        private int _bits;

        public BitWriter(Stream stream) {
            _stream = stream;
        }

        public void WriteBits(uint bits, int count) {
            _buffer = (_buffer << count) | (bits & ((1u << count) - 1));
            _bits += count;
            while (_bits >= 8) {
                var b = (byte)((_buffer >> (_bits - 8)) & 0xFF);
                WriteByte(b);
                _bits -= 8;
            }
        }

        public void Flush() {
            if (_bits > 0) {
                int padding = 8 - _bits;
                var b = (byte)((_buffer << padding) | ((1U << padding) - 1U));
                WriteByte(b);
                _bits = 0;
            }
            FlushBytes();
        }

        private void WriteByte(byte b) {
            AppendByte(b);
            if (b == 0xFF) AppendByte(0x00);
        }

        private void AppendByte(byte value) {
            _bytes[_byteCount++] = value;
            if (_byteCount == _bytes.Length) FlushBytes();
        }

        private void FlushBytes() {
            if (_byteCount == 0) return;
            _stream.Write(_bytes, 0, _byteCount);
            _byteCount = 0;
        }
    }
}

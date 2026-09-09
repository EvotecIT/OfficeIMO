using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Bounded T.4/T.6 decoding shared by document and raster readers.</summary>
internal static partial class OfficeFaxDecoder {
    internal static byte[] Decode(byte[] encoded, int columns, int rows, int k, bool endOfLine,
        bool byteAligned, bool blackIsOne, bool endOfBlock, int maximumBytes, CancellationToken cancellationToken) {
        if (columns <= 0 || rows <= 0) throw new InvalidDataException("Fax decoding requires positive image dimensions.");
        long strideLong = ((long)columns + 7) / 8;
        long length = strideLong * rows;
        if (length > maximumBytes) throw new InvalidDataException("Fax image exceeds the decoded byte limit.");
        int stride = (int)strideLong;
        var output = new byte[(int)length];
        var bits = new FaxBits(encoded, cancellationToken);
        for (int row = 0; row < rows; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            bool foundEndOfLine = bits.TryReadEndOfLine();
            if (endOfLine && !foundEndOfLine) throw new InvalidDataException("Missing fax end-of-line marker.");
            bool oneDimensional = k == 0 || (k > 0 && bits.Read() != 0);
            if (oneDimensional) DecodeOneDimensional(bits, output, row * stride, columns, cancellationToken);
            else DecodeTwoDimensional(bits, output, row * stride, row == 0 ? -1 : (row - 1) * stride, columns, cancellationToken);
            // Group 3 fill bits precede the next EOL (and its optional 2-D tag).
            // Let TryReadEndOfLine consume them; aligning here can skip into the marker.
            if (byteAligned && !foundEndOfLine && !endOfLine) bits.Align();
        }
        if (endOfBlock) {
            int markers = k < 0 ? 2 : 6;
            for (int marker = 0; marker < markers; marker++) {
                if (!bits.TryReadEndOfLine() || (k > 0 && bits.Read() != 1)) {
                    throw new InvalidDataException("Missing fax end-of-block marker.");
                }
            }
        }
        if (!blackIsOne) {
            for (int index = 0; index < output.Length; index++) {
                if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                output[index] = (byte)~output[index];
            }
        }
        return output;
    }

    private static void DecodeOneDimensional(FaxBits bits, byte[] output, int offset, int columns, CancellationToken token) {
        int x = 0;
        bool black = false;
        while (x < columns) {
            token.ThrowIfCancellationRequested();
            int run = ReadRun(bits, black, columns - x);
            Paint(output, offset, x, x + run, black);
            x += run;
            black = !black;
        }
    }

    private static void DecodeTwoDimensional(FaxBits bits, byte[] output, int offset, int reference, int columns, CancellationToken token) {
        int x = 0;
        bool black = false;
        bool first = true;
        while (x < columns) {
            token.ThrowIfCancellationRequested();
            int mode = ReadMode(bits);
            if (mode == 10) {
                int middle = x + ReadRun(bits, black, columns - x);
                int end = middle + ReadRun(bits, !black, columns - middle);
                Paint(output, offset, x, middle, black);
                Paint(output, offset, middle, end, !black);
                if (end == x) throw new InvalidDataException("Fax horizontal mode made no progress.");
                x = end;
            } else {
                int b1 = FindReferenceChange(output, reference, columns, x, black, first);
                if (mode == 11) {
                    int b2 = b1 < columns ? FindReferenceChange(output, reference, columns, b1, !black, false) : columns;
                    if (b2 <= x) throw new InvalidDataException("Invalid fax pass mode.");
                    Paint(output, offset, x, b2, black);
                    x = b2;
                } else {
                    int end = b1 + mode;
                    if (end < x || end > columns) throw new InvalidDataException("Fax vertical run is outside its row.");
                    Paint(output, offset, x, end, black);
                    x = end;
                    black = !black;
                }
            }
            first = false;
        }
    }

    private static int FindReferenceChange(byte[] output, int reference, int columns, int x, bool black, bool first) {
        if (reference < 0) return columns;
        int start = first ? x : x + 1;
        bool previous = start > 0 && IsBlack(output, reference, start - 1);
        for (int position = start; position < columns; position++) {
            bool current = IsBlack(output, reference, position);
            if (current != previous && current != black) return position;
            previous = current;
        }
        return columns;
    }

    private static bool IsBlack(byte[] output, int offset, int x) => (output[offset + x / 8] & (128 >> (x & 7))) != 0;

    private static void Paint(byte[] output, int offset, int start, int end, bool black) {
        if (!black) return;
        while (start < end && (start & 7) != 0) { output[offset + start / 8] |= (byte)(128 >> (start & 7)); start++; }
        while (end - start >= 8) { output[offset + start / 8] = 255; start += 8; }
        while (start < end) { output[offset + start / 8] |= (byte)(128 >> (start & 7)); start++; }
    }

    private static int ReadRun(FaxBits bits, bool black, int remaining) {
        Dictionary<int, int> codes = black ? BlackCodes : WhiteCodes;
        int total = 0;
        while (true) {
            int code = 1;
            int run = -1;
            for (int length = 1; length <= 13; length++) {
                code = (code << 1) | bits.Read();
                if (codes.TryGetValue(code, out run)) break;
                run = -1;
            }
            if (run < 0 || run > remaining - total) throw new InvalidDataException("Invalid fax run length.");
            total += run;
            if (run < 64) return total;
        }
    }

    private static int ReadMode(FaxBits bits) {
        int code = 1;
        for (int length = 1; length <= 7; length++) {
            code = (code << 1) | bits.Read();
            switch (code) {
                case 3: return 0; // 1
                case 9: return 10; // 001: horizontal
                case 10: return -1; // 010
                case 11: return 1; // 011
                case 17: return 11; // 0001: pass
                case 66: return -2; // 000010
                case 67: return 2; // 000011
                case 130: return -3; // 0000010
                case 131: return 3; // 0000011
            }
        }
        throw new InvalidDataException("Unsupported or invalid fax two-dimensional mode.");
    }

    private sealed class FaxBits {
        private readonly byte[] _bytes;
        private readonly CancellationToken _token;
        private long _position;
        internal FaxBits(byte[] bytes, CancellationToken token) { _bytes = bytes; _token = token; }
        internal int Read() {
            if (_position >= (long)_bytes.Length * 8) throw new InvalidDataException("Truncated fax image.");
            int value = (_bytes[(int)(_position / 8)] >> (7 - (int)(_position & 7))) & 1;
            _position++;
            return value;
        }
        internal bool TryReadEndOfLine() {
            long saved = _position;
            int zeros = 0;
            while (_position < (long)_bytes.Length * 8) {
                if ((_position & 4095) == 0) _token.ThrowIfCancellationRequested();
                if (Read() != 0) {
                    if (zeros >= 11) return true;
                    break;
                }
                // T.4 fill has variable length. Saturate the marker threshold so a long
                // bounded payload cannot overflow the counter, while cancellation stays responsive.
                if (zeros < 11) zeros++;
            }
            _position = saved;
            return false;
        }
        internal void Align() { _position = (_position + 7) & ~7L; }
    }
}

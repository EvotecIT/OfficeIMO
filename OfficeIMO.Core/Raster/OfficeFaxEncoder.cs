using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Lossless T.6 Group 4 encoding of opaque black/white RGBA pixels, with bounded row and output storage.</summary>
internal static class OfficeFaxEncoder {
    private static readonly int[] WhiteCodes = OfficeFaxDecoder.CreateRunEncodingTable(black: false);
    private static readonly int[] BlackCodes = OfficeFaxDecoder.CreateRunEncodingTable(black: true);

    internal static byte[] EncodeGroup4(OfficeRasterImage image, CancellationToken token = default,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null) {
        token.ThrowIfCancellationRequested();
        // Two transition rows avoid rescanning a long reference run for every short coding run.
        // The caller includes these arrays in its compression working-set reservation.
        var current = new int[checked(image.Width + 2)];
        var reference = new int[current.Length];
        var counter = new FaxWriter(null);
        EncodeRows(image, current, reference, ref counter, token, checkpointObserver);
        byte[] output = new byte[counter.Length];
        var writer = new FaxWriter(output);
        EncodeRows(image, current, reference, ref writer, token, checkpointObserver);
        if (writer.Length != output.Length) throw new InvalidOperationException("Fax strip length changed during encoding.");
        return output;
    }

    private static void EncodeRows(OfficeRasterImage image, int[] current, int[] reference,
        ref FaxWriter writer, CancellationToken token, Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        int width = image.Width;
        reference[0] = reference[1] = width; // The imaginary first reference row is white.
        int referenceCount = 2;
        for (int row = 0; row < image.Height; row++) {
            checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.TiffCompressionRow);
            token.ThrowIfCancellationRequested();
            int currentCount = ReadTransitions(image.PixelBuffer, row * width * 4, width, current, token);
            int x = 0, aIndex = 0, referenceIndex = 0;
            bool black = false, first = true;
            while (x < width) {
                token.ThrowIfCancellationRequested();
                int a1 = current[aIndex];
                // b1 is the first reference transition to the opposite color to the right of a0.
                // Only at the start of a row may a transition at column zero be selected.
                while (referenceIndex < referenceCount &&
                    (first ? reference[referenceIndex] < x : reference[referenceIndex] <= x)) referenceIndex++;
                int bIndex = referenceIndex;
                if ((bIndex & 1) != (black ? 1 : 0)) bIndex++;
                int b1 = bIndex < referenceCount ? reference[bIndex] : width;
                int b2 = bIndex + 1 < referenceCount ? reference[bIndex + 1] : width;
                if (b2 < a1) {
                    writer.Write(1, 4); // Pass: 0001.
                    x = b2;
                } else {
                    int delta = a1 - b1;
                    if (delta >= -3 && delta <= 3) {
                        if (delta == 0) writer.Write(1, 1);
                        else writer.Write(delta < 0 ? 2 : 3, Math.Abs(delta) == 1 ? 3 : Math.Abs(delta) + 4);
                        x = a1;
                        aIndex++;
                        black = !black;
                    } else {
                        int a2 = current[aIndex + 1];
                        writer.Write(1, 3); // Horizontal: 001, followed by two T.4 run codes.
                        WriteRun(ref writer, a1 - x, black);
                        WriteRun(ref writer, a2 - a1, !black);
                        x = a2;
                        aIndex += 2;
                    }
                }
                first = false;
            }
            int[] swap = reference; reference = current; current = swap;
            referenceCount = currentCount;
        }
        writer.Write(1, 12);
        writer.Write(1, 12); // EOFB: two EOL codewords, then zero padding to a byte boundary.
        token.ThrowIfCancellationRequested();
    }

    private static int ReadTransitions(byte[] pixels, int offset, int width, int[] transitions, CancellationToken token) {
        int count = 0;
        bool previousBlack = false;
        for (int x = 0; x < width; x++, offset += 4) {
            if ((x & 4095) == 0) token.ThrowIfCancellationRequested();
            byte sample = pixels[offset];
            if ((sample != 0 && sample != 255) || pixels[offset + 1] != sample ||
                pixels[offset + 2] != sample || pixels[offset + 3] != 255) {
                throw new ArgumentException("CCITT Group 4 requires opaque pixels whose RGB samples are all zero or all 255. Convert to bilevel explicitly before encoding.", nameof(pixels));
            }
            bool black = sample == 0;
            if (black != previousBlack) transitions[count++] = x;
            previousBlack = black;
        }
        transitions[count++] = width;
        transitions[count++] = width;
        return count;
    }

    private static void WriteRun(ref FaxWriter writer, int run, bool black) {
        int[] codes = black ? BlackCodes : WhiteCodes;
        while (run >= 2560) { writer.WriteCode(codes[103]); run -= 2560; }
        if (run >= 64) { writer.WriteCode(codes[63 + run / 64]); run %= 64; }
        writer.WriteCode(codes[run]); // A terminating code is required even after an exact makeup run.
    }

    private struct FaxWriter {
        private readonly byte[]? _output;
        private long _position;
        internal FaxWriter(byte[]? output) { _output = output; _position = 0; }
        internal int Length => checked((int)((_position + 7) / 8));

        internal void WriteCode(int code) {
            int length = 0;
            for (int sentinel = code; sentinel > 1; sentinel >>= 1) length++;
            if (length == 0) throw new InvalidOperationException("Missing fax run code.");
            Write(code, length);
        }

        internal void Write(int value, int length) {
            if (_position + length > (long)OfficeRasterGuards.MaximumEncodedBytes * 8L)
                throw new ArgumentException("The fax strip exceeds the encoded-size limit.");
            if (_output != null) {
                for (int bit = length - 1; bit >= 0; bit--, _position++) {
                    int index = (int)(_position / 8);
                    if (index >= _output.Length) throw new InvalidOperationException("Fax pixels changed during encoding.");
                    if (((value >> bit) & 1) != 0) _output[index] |= (byte)(128 >> (int)(_position & 7));
                }
            } else _position += length;
        }
    }
}

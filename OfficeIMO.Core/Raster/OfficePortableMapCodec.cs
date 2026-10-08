using System;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Managed PBM, PGM, and PPM decoding and bilevel PBM writing.</summary>
internal static class OfficePortableMapCodec {
    internal static bool TryIdentify(byte[] bytes, out int width, out int height, CancellationToken token = default) {
        width = height = 0;
        token.ThrowIfCancellationRequested();
        try { if (bytes.Length < 7 || bytes[0] != 'P' || bytes[1] < '1' || bytes[1] > '6' || !White(bytes[2])) return false; int cursor = 2; width = Number(bytes, ref cursor, token); height = Number(bytes, ref cursor, token); return OfficeRasterGuards.TryEnsurePixelCount(width, height, out _); }
        catch (Exception ex) when (ex is FormatException || ex is OverflowException) { return false; }
    }
    internal static bool TryDecode(byte[] bytes, OfficeRasterDecodeOptions options, out OfficeRasterImage? image) {
        image = null;
        CancellationToken token = options.CancellationToken;
        token.ThrowIfCancellationRequested();
        try {
            if (bytes.Length < 7 || bytes[0] != 'P' || bytes[1] < '1' || bytes[1] > '6' || !White(bytes[2])) return false;
            int kind = bytes[1] - '0'; int cursor = 2;
            int width = Number(bytes, ref cursor, token); int height = Number(bytes, ref cursor, token);
            if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, options.MaximumDecodedPixels, out int count) || (long)count * 4 + bytes.LongLength + options.RetainedManagedBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
            int maximum = kind == 1 || kind == 4 ? 1 : Number(bytes, ref cursor, token); if (maximum < 1 || maximum > 65535) return false;
            if (kind >= 4) {
                // A comment may immediately follow the final header number. Its line
                // ending supplies the raster delimiter; subsequent bytes are pixels.
                if (cursor < bytes.Length && bytes[cursor] == '#') {
                    SkipComment(bytes, ref cursor, token);
                }
                if (cursor >= bytes.Length || !White(bytes[cursor])) return false;
                if (bytes[cursor++] == '\r' && cursor < bytes.Length && bytes[cursor] == '\n') cursor++;
            }
            int samples = kind == 3 || kind == 6 ? 3 : 1;
            int sampleBytes = maximum > 255 ? 2 : 1;
            long expected = kind == 4 ? ((long)width + 7) / 8 * height : (long)count * samples * sampleBytes;
            if (kind >= 4 && expected != bytes.LongLength - cursor) return false;
            token.ThrowIfCancellationRequested();
            var result = new OfficeRasterImage(width, height);
            int rowBytes = (width + 7) / 8;
            for (int y = 0; y < height; y++) {
                for (int x = 0; x < width; x++) {
                    // A valid image may put the entire pixel budget in one row.
                    if ((x & 4095) == 0) token.ThrowIfCancellationRequested();
                    int red, green, blue;
                    if (kind == 4) { red = ((bytes[cursor + y * rowBytes + x / 8] >> (7 - x % 8)) & 1) != 0 ? 0 : 255; green = blue = red; }
                    else if (kind == 1) {
                        Skip(bytes, ref cursor, token);
                        if (cursor >= bytes.Length || bytes[cursor] != '0' && bytes[cursor] != '1') return false;
                        red = bytes[cursor++] == '1' ? 0 : 255;
                        green = blue = red;
                    }
                    else {
                        int first = Sample(); red = (first * 255 + maximum / 2) / maximum;
                        if (samples == 3) { green = (Sample() * 255 + maximum / 2) / maximum; blue = (Sample() * 255 + maximum / 2) / maximum; } else green = blue = red;
                    }
                    result.SetPixel(x, y, OfficeColor.FromRgb((byte)red, (byte)green, (byte)blue));
                }
            }
            if (kind < 4) { Skip(bytes, ref cursor, token); if (cursor != bytes.Length) return false; }
            token.ThrowIfCancellationRequested();
            image = result; return true;
            int Sample() {
                int value = kind < 4 ? Number(bytes, ref cursor, token) : sampleBytes == 1 ? bytes[cursor++] : bytes[cursor++] << 8 | bytes[cursor++];
                if (value < 0 || value > maximum) throw new FormatException("Portable map sample is outside its declared range."); return value;
            }
        } catch (OperationCanceledException) { throw; } catch (Exception ex) when (ex is FormatException || ex is ArgumentException || ex is OverflowException || ex is IndexOutOfRangeException) { return false; }
    }
    internal static void EncodePbmTo(OfficeRasterImage image, Stream output, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        OfficeRasterGuards.EnsureOutputPixels(image.Width, image.Height, "PBM dimensions exceed pixel limits.");
        int rowLength = (image.Width + 7) / 8;
        byte[] header = Encoding.ASCII.GetBytes("P4\n" + image.Width + " " + image.Height + "\n");
        int size = OfficeRasterGuards.EnsureOutputBytes(header.Length + rowLength * (long)image.Height, "PBM output exceeds the encoded-size limit.");
        OfficeRasterOutput.EnsureImageWriteWorkingSet(image, output, size, rowLength, header.Length,
            "PBM encoding exceeds the managed working-set limit.");
        output.Write(header, 0, header.Length);
        token.ThrowIfCancellationRequested();
        byte[] row = new byte[rowLength];
        for (int y = 0; y < image.Height; y++) {
            token.ThrowIfCancellationRequested(); Array.Clear(row, 0, row.Length);
            for (int x = 0; x < image.Width; x++) {
                if ((x & 4095) == 0) token.ThrowIfCancellationRequested();
                OfficeColor color = image.GetPixel(x, y);
                double luminance = (0.2126 * color.R + 0.7152 * color.G + 0.0722 * color.B) * color.A / 255D + 255D - color.A;
                if (luminance < 128) row[x / 8] |= (byte)(1 << (7 - x % 8));
            }
            token.ThrowIfCancellationRequested();
            output.Write(row, 0, row.Length);
        }
        token.ThrowIfCancellationRequested();
    }
    private static int Number(byte[] bytes, ref int cursor, CancellationToken token) {
        Skip(bytes, ref cursor, token); int value = 0; int start = cursor;
        while (cursor < bytes.Length && bytes[cursor] >= '0' && bytes[cursor] <= '9') {
            if ((cursor & 4095) == 0) token.ThrowIfCancellationRequested();
            value = checked(value * 10 + bytes[cursor++] - '0');
        }
        token.ThrowIfCancellationRequested();
        if (cursor == start || cursor < bytes.Length && !White(bytes[cursor]) && bytes[cursor] != '#') throw new FormatException("Invalid portable map number."); return value;
    }
    private static void Skip(byte[] bytes, ref int cursor, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        while (cursor < bytes.Length) {
            if ((cursor & 4095) == 0) token.ThrowIfCancellationRequested();
            if (White(bytes[cursor])) cursor++;
            else if (bytes[cursor] == '#') SkipComment(bytes, ref cursor, token);
            else break;
        }
        token.ThrowIfCancellationRequested();
    }
    private static void SkipComment(byte[] bytes, ref int cursor, CancellationToken token) {
        while (cursor < bytes.Length && bytes[cursor] != '\n' && bytes[cursor] != '\r') {
            if ((cursor & 4095) == 0) token.ThrowIfCancellationRequested();
            cursor++;
        }
        token.ThrowIfCancellationRequested();
    }
    private static bool White(byte value) => value == 9 || value == 10 || value == 11 || value == 12 || value == 13 || value == 32;
}

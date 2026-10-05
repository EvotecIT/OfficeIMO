using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    private static bool TryDecodeJpegSegments(byte[] bytes, IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian, int width, int height, int samples, int photometric, int planar,
        OfficeRasterDecodeOptions options, TiffValidationBudget? budget, bool retainPixels, out byte[] source) {
        source = Array.Empty<byte>();
        bool ycc = photometric == 6;
        int horizontal = 1, vertical = 1, positioning = 1;
        double[] coefficients = { .299, .587, .114 }, reference = { 0, 255, 128, 255, 128, 255 };
        if (ycc) {
            if (!TryReadScalarOrDefault(bytes, entries, 531, littleEndian, 1, out positioning) || (positioning != 1 && positioning != 2)) return false;
            if (entries.ContainsKey(530)) {
                if (!TryReadValues(bytes, entries, 530, littleEndian, 2, out int[] subsampling)) return false;
                horizontal = subsampling[0]; vertical = subsampling[1];
            } else horizontal = vertical = 2;
            if ((horizontal != 1 && horizontal != 2 && horizontal != 4) ||
                (vertical != 1 && vertical != 2 && vertical != 4) || vertical > horizontal ||
                !TryReadJpegRationals(bytes, entries, 529, littleEndian, coefficients) ||
                !TryReadJpegRationals(bytes, entries, 532, littleEndian, reference) ||
                coefficients[0] <= 0 || coefficients[1] <= 0 || coefficients[2] <= 0 ||
                Math.Abs(coefficients[0] + coefficients[1] + coefficients[2] - 1) > .00001 ||
                reference[1] <= reference[0] || reference[3] <= reference[2] || reference[5] <= reference[4]) return false;
        }
        bool strips = entries.ContainsKey(273) || entries.ContainsKey(279);
        bool tiles = entries.ContainsKey(324) || entries.ContainsKey(325) || entries.ContainsKey(322) || entries.ContainsKey(323);
        if (strips == tiles) return false;
        int sw = width, sh;
        if (strips) {
            if (!TryReadRowsPerStrip(bytes, entries, littleEndian, height, out sh)) return false;
        } else if (!TryReadScalar(bytes, entries, 322, littleEndian, out sw) ||
            !TryReadScalar(bytes, entries, 323, littleEndian, out sh) || sw < 1 || sh < 1) return false;
        if (ycc && (strips ? sh < height && sh % vertical != 0 : sw % horizontal != 0 || sh % vertical != 0)) return false;
        int across = checked((int)(((long)width + sw - 1) / sw));
        int down = checked((int)(((long)height + sh - 1) / sh));
        int perPlane = checked(across * down), count = checked(perPlane * (planar == 2 ? samples : 1));
        int sourceLength = OfficeRasterGuards.EnsureByteCount((long)width * height * samples, "TIFF JPEG pixels exceed the managed limit.");
        long retained = checked(options.RetainedManagedBytes + bytes.LongLength + (long)count * 8 +
            (retainPixels ? sourceLength + (long)width * height * 4 : 0) + 65536);
        if (retained > OfficeRasterGuards.MaximumDecodedBytes ||
            !TryReadValues(bytes, entries, strips ? 273 : 324, littleEndian, count, options.CancellationToken, out int[] offsets) ||
            !TryReadValues(bytes, entries, strips ? 279 : 325, littleEndian, count, options.CancellationToken, out int[] lengths)) return false;
        byte[] tables = Array.Empty<byte>();
        int inherited = 0;
        if (entries.TryGetValue(347, out TiffEntry tableEntry)) {
            if (tableEntry.Type != 7 || tableEntry.Count < 4 || retained + tableEntry.Count > OfficeRasterGuards.MaximumDecodedBytes) return false;
            int tableOffset = tableEntry.Count <= 4 ? tableEntry.ValueFieldOffset : ReadOffset(bytes, tableEntry.ValueFieldOffset, littleEndian);
            if (!HasBytes(bytes, tableOffset, tableEntry.Count)) return false;
            tables = new byte[tableEntry.Count];
            CopyWithCancellation(bytes, tableOffset, tables, 0, tables.Length, options.CancellationToken);
            if (!TryNormalizeTiffJpeg(tables, true, 0, 0, 0, 1, 1, 0, options.CancellationToken, out inherited, out _)) return false;
            retained += tables.Length;
        }
        if (retainPixels) source = new byte[sourceLength];
        int frameProcess = 0;
        for (int segment = 0; segment < count; segment++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            int tile = segment % perPlane, plane = segment / perPlane;
            int left = tile % across * sw, top = tile / across * sh;
            int rows = Math.Min(sh, height - top), columns = Math.Min(sw, width - left);
            int decodeWidth = sw, decodeRows = strips ? rows : sh, channels = planar == 2 ? 1 : samples;
            if (ycc && planar == 2 && plane > 0) {
                decodeWidth = checked((int)(((long)sw + horizontal - 1) / horizontal));
                decodeRows = checked((int)(((long)decodeRows + vertical - 1) / vertical));
            }
            int expected = OfficeRasterGuards.EnsureByteCount((long)decodeWidth * decodeRows * channels, "TIFF JPEG segment exceeds the managed limit.");
            if (!HasSegment(bytes, offsets[segment], lengths[segment]) || lengths[segment] < 4 ||
                budget != null && !budget.TryReserve(checked(lengths[segment] + tables.Length), expected)) return false;
            int combinedLength = checked(lengths[segment] + (tables.Length == 0 ? 0 : tables.Length - 4));
            // TIFF reconstruction retains two reduced chroma planes when positioning
            // differs from JPEG or a partial tile needs image-boundary clamping.
            bool reconstructChroma = ycc && planar == 1 && (positioning == 2 || columns < decodeWidth || rows < decodeRows);
            long chromaScratch = reconstructChroma ? expected : 0;
            // The temporary segment copy and combined stream coexist with the TIFF and output.
            if (retained + lengths[segment] + combinedLength + expected + chromaScratch > OfficeRasterGuards.MaximumDecodedBytes) return false;
            var jpeg = new byte[lengths[segment]];
            CopyWithCancellation(bytes, offsets[segment], jpeg, 0, jpeg.Length, options.CancellationToken);
            if (!TryNormalizeTiffJpeg(jpeg, false, decodeWidth, decodeRows, channels, ycc && planar == 1 ? horizontal : 1,
                ycc && planar == 1 ? vertical : 1, inherited, options.CancellationToken, out _, out int segmentProcess)) return false;
            if (frameProcess != 0 && segmentProcess != frameProcess) return false;
            frameProcess = segmentProcess;
            byte[] combined = jpeg;
            if (tables.Length > 0) {
                combined = new byte[combinedLength];
                CopyWithCancellation(tables, 0, combined, 0, tables.Length - 2, options.CancellationToken);
                CopyWithCancellation(jpeg, 2, combined, tables.Length - 2, jpeg.Length - 2, options.CancellationToken);
            }
            if (!OfficeJpegCodec.TryDecodeColorComponents(combined, 0, false, out byte[] decoded,
                out int jw, out int jh, out int jc, options: new OfficeJpegDecodeOptions(highQualityChroma: ycc && !reconstructChroma), cancellationToken: options.CancellationToken,
                retainedManagedBytes: retained + jpeg.Length + chromaScratch) || jw != decodeWidth || jh != decodeRows || jc != channels || decoded.Length != expected) return false;
            if (!retainPixels) continue;
            if (ycc && planar == 1) {
                if (reconstructChroma) ReconstructTiffJpegChroma(decoded, decodeWidth, columns, rows, horizontal, vertical, positioning, options);
                ConvertTiffJpegYcc(decoded, coefficients, reference, options);
            }
            if (planar == 2) {
                if (ycc && plane > 0) CopyTiffJpegChroma(decoded, source, plane, width, left, top, columns, rows,
                    decodeWidth, decodeRows, horizontal, vertical, positioning, options);
                else if (strips) CopyPlanarRows(decoded, source, plane, samples, 1, width, top, rows, options);
                else CopyTile(decoded, source, plane, planar, samples, 1, width, height, left, top, sw, sh, options);
            } else for (int row = 0; row < rows; row++)
                CopyWithCancellation(decoded, row * sw * samples, source, ((top + row) * width + left) * samples,
                    columns * samples, options.CancellationToken);
        }
        if (ycc && planar == 2 && retainPixels) ConvertTiffJpegYcc(source, coefficients, reference, options);
        return true;
    }

    private static bool TryReadJpegRationals(byte[] bytes, IReadOnlyDictionary<int, TiffEntry> entries,
        int tag, bool littleEndian, double[] values) {
        if (!entries.TryGetValue(tag, out TiffEntry entry)) return true;
        if (entry.Type != 5 || entry.Count != values.Length) return false;
        int offset = ReadOffset(bytes, entry.ValueFieldOffset, littleEndian);
        if (!HasBytes(bytes, offset, values.Length * 8)) return false;
        for (int i = 0; i < values.Length; i++) {
            uint denominator = ReadUInt32(bytes, offset + i * 8 + 4, littleEndian);
            if (denominator == 0) return false;
            values[i] = ReadUInt32(bytes, offset + i * 8, littleEndian) / (double)denominator;
        }
        return true;
    }

    private static void ConvertTiffJpegYcc(byte[] pixels, double[] c, double[] r, OfficeRasterDecodeOptions options) {
        byte Clamp(double value) => (byte)Math.Round(Math.Max(0, Math.Min(255, value)));
        for (int i = 0; i < pixels.Length; i += 3) {
            if ((i & 4095) == 0) options.CancellationToken.ThrowIfCancellationRequested();
            double y = (pixels[i] - r[0]) * 255 / (r[1] - r[0]);
            double cb = (pixels[i + 1] - r[2]) * 127 / (r[3] - r[2]);
            double cr = (pixels[i + 2] - r[4]) * 127 / (r[5] - r[4]);
            double red = y + cr * (2 - 2 * c[0]), blue = y + cb * (2 - 2 * c[2]);
            pixels[i] = Clamp(red); pixels[i + 1] = Clamp((y - c[0] * red - c[2] * blue) / c[1]); pixels[i + 2] = Clamp(blue);
        }
    }
}

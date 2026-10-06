using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    private static bool TryDecodeJpegSegments(byte[] bytes, IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian, int width, int height, int samples, int sampleBytes, int sampleBits, int photometric, int planar,
        OfficeRasterDecodeOptions options, TiffValidationBudget? budget, bool retainPixels, out byte[] source, out TiffJpegColorTransform? jpegColor, bool legacy = false) {
        source = Array.Empty<byte>();
        jpegColor = null;
        int interchangeOffset = 0, interchangeLength = 0, legacyProcess = 1;
        if (legacy && (!TryReadScalarOrDefault(bytes, entries, 512, littleEndian, 1, out legacyProcess) ||
            (legacyProcess != 1 && legacyProcess != 14) || (legacyProcess == 1 && sampleBits != 8 && sampleBits != 12) ||
            !TryGetLegacyJpegInterchange(bytes, entries, littleEndian, out interchangeOffset, out interchangeLength))) return false;
        bool fullInterchange = interchangeLength > 0;
        if (fullInterchange) planar = 1; // The interchange stream describes the complete image.
        bool ycc = photometric == 6;
        int horizontal = 1, vertical = 1, positioning = 1;
        int maximum = (1 << sampleBits) - 1, midpoint = (maximum + 1) / 2;
        double[] coefficients = { .299, .587, .114 }, reference = { 0, maximum, midpoint, maximum, midpoint, maximum };
        if (ycc) {
            if (!TryReadScalarOrDefault(bytes, entries, 531, littleEndian, 1, out positioning) || (positioning != 1 && positioning != 2)) return false;
            if (entries.ContainsKey(530)) {
                if (!TryReadValues(bytes, entries, 530, littleEndian, 2, out int[] subsampling)) return false;
                horizontal = subsampling[0]; vertical = subsampling[1];
            } else horizontal = vertical = 2;
            if ((horizontal != 1 && horizontal != 2 && horizontal != 4) ||
                (vertical != 1 && vertical != 2 && vertical != 4) || vertical > horizontal ||
                !TryReadJpegRationals(bytes, entries, 529, littleEndian, coefficients) ||
                !TryReadJpegRationals(bytes, entries, 532, littleEndian, reference, allowInteger: legacy) ||
                coefficients[0] <= 0 || coefficients[1] <= 0 || coefficients[2] <= 0 ||
                Math.Abs(coefficients[0] + coefficients[1] + coefficients[2] - 1) > .00001 ||
                reference[1] <= reference[0] || reference[3] <= reference[2] || reference[5] <= reference[4]) return false;
        }
        bool strips = fullInterchange || entries.ContainsKey(273) || entries.ContainsKey(279);
        bool tiles = !fullInterchange && (entries.ContainsKey(324) || entries.ContainsKey(325) || entries.ContainsKey(322) || entries.ContainsKey(323));
        if (strips == tiles) return false;
        int sw = width, sh;
        if (strips) {
            if (fullInterchange) sh = height;
            else if (!TryReadRowsPerStrip(bytes, entries, littleEndian, height, out sh)) return false;
        } else if (!TryReadScalar(bytes, entries, 322, littleEndian, out sw) ||
            !TryReadScalar(bytes, entries, 323, littleEndian, out sh) || sw < 1 || sh < 1) return false;
        if (ycc && (strips ? sh < height && sh % vertical != 0 : sw % horizontal != 0 || sh % vertical != 0)) return false;
        int across = checked((int)(((long)width + sw - 1) / sw));
        int down = checked((int)(((long)height + sh - 1) / sh));
        int perPlane = checked(across * down), count = checked(perPlane * (planar == 2 ? samples : 1));
        int sourceLength = OfficeRasterGuards.EnsureByteCount((long)width * height * samples * sampleBytes, "TIFF JPEG pixels exceed the managed limit.");
        int fractionLength = ycc && retainPixels && (horizontal > 1 || vertical > 1)
            ? OfficeRasterGuards.EnsureByteCount((long)width * height * 2, "TIFF chroma fractions exceed the managed limit.") : 0;
        long retained = checked(options.RetainedManagedBytes + bytes.LongLength + (long)count * 8 +
            (retainPixels ? sourceLength + (long)width * height * 4 + fractionLength : 0) + 65536);
        if (retained > OfficeRasterGuards.MaximumDecodedBytes) return false;
        byte[]? chromaFractions = fractionLength == 0 ? null : new byte[fractionLength];
        if (ycc) jpegColor = new TiffJpegColorTransform(maximum, coefficients, reference, chromaFractions);
        int[] offsets, lengths;
        if (fullInterchange) { offsets = new[] { interchangeOffset }; lengths = new[] { interchangeLength }; }
        else if (!TryReadValues(bytes, entries, strips ? 273 : 324, littleEndian, count, options.CancellationToken, out offsets) ||
            !TryReadValues(bytes, entries, strips ? 279 : 325, littleEndian, count, options.CancellationToken, out lengths)) return false;
        byte[] tables = Array.Empty<byte>();
        int inherited = 0;
        if (!legacy && entries.TryGetValue(347, out TiffEntry tableEntry)) {
            if (tableEntry.Type != 7 || tableEntry.Count < 4 || retained + tableEntry.Count > OfficeRasterGuards.MaximumDecodedBytes) return false;
            int tableOffset = tableEntry.Count <= 4 ? tableEntry.ValueFieldOffset : ReadOffset(bytes, tableEntry.ValueFieldOffset, littleEndian);
            if (!HasBytes(bytes, tableOffset, tableEntry.Count)) return false;
            tables = new byte[tableEntry.Count];
            CopyWithCancellation(bytes, tableOffset, tables, 0, tables.Length, options.CancellationToken);
            if (!TryNormalizeTiffJpeg(tables, true, 0, 0, 0, sampleBits, 1, 1, 0, options.CancellationToken, out inherited, out _)) return false;
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
            if (ycc && planar == 2 && (plane == 1 || plane == 2)) {
                decodeWidth = checked((int)(((long)sw + horizontal - 1) / horizontal));
                decodeRows = checked((int)(((long)decodeRows + vertical - 1) / vertical));
            }
            int expected = OfficeRasterGuards.EnsureByteCount((long)decodeWidth * decodeRows * channels * sampleBytes, "TIFF JPEG segment exceeds the managed limit.");
            if (!HasSegment(bytes, offsets[segment], lengths[segment]) || lengths[segment] < (legacy ? 1 : 4) ||
                budget != null && !budget.TryReserve(checked(lengths[segment] + tables.Length + (legacy ? LegacyJpegHeaderLimit : 0)), expected)) return false;
            int segmentLimit = checked(lengths[segment] + (legacy ? LegacyJpegHeaderLimit : 0));
            int combinedLength = checked(segmentLimit + (tables.Length == 0 ? 0 : tables.Length - 4));
            // Decode nearest native samples so TIFF can retain exact interpolation
            // fractions for both centered/cosited grids and partial edge tiles.
            bool reconstructChroma = ycc && planar == 1 && (horizontal > 1 || vertical > 1);
            long chromaScratch = reconstructChroma ? expected : 0;
            // The temporary segment copy and combined stream coexist with the TIFF and output.
            if (retained + segmentLimit + combinedLength + expected + chromaScratch > OfficeRasterGuards.MaximumDecodedBytes) return false;
            byte[] jpeg;
            if (legacy) {
                if (!TryReconstructLegacyJpeg(bytes, offsets[segment], lengths[segment], entries, littleEndian,
                    decodeWidth, decodeRows, samples, channels, planar == 2 ? plane : -1, sampleBits,
                    legacyProcess, ycc && planar == 1 ? horizontal : 1, ycc && planar == 1 ? vertical : 1,
                    options, out jpeg)) return false;
            } else {
                jpeg = new byte[lengths[segment]];
                CopyWithCancellation(bytes, offsets[segment], jpeg, 0, jpeg.Length, options.CancellationToken);
            }
            if (!TryNormalizeTiffJpeg(jpeg, false, decodeWidth, decodeRows, channels, sampleBits, ycc && planar == 1 ? horizontal : 1,
                ycc && planar == 1 ? vertical : 1, inherited, options.CancellationToken, out _, out int segmentProcess)) return false;
            if (legacy && (legacyProcess == 14 ? segmentProcess != 195 : segmentProcess != 192 && segmentProcess != 193)) return false;
            if (frameProcess != 0 && segmentProcess != frameProcess) return false;
            frameProcess = segmentProcess;
            byte[] combined = jpeg;
            if (tables.Length > 0) {
                combined = new byte[checked(jpeg.Length + tables.Length - 4)];
                CopyWithCancellation(tables, 0, combined, 0, tables.Length - 2, options.CancellationToken);
                CopyWithCancellation(jpeg, 2, combined, tables.Length - 2, jpeg.Length - 2, options.CancellationToken);
            }
            if (!OfficeJpegCodec.TryDecodeColorComponents(combined, 0, false, out byte[] decoded,
                out int jw, out int jh, out int jc, options: new OfficeJpegDecodeOptions(highQualityChroma: ycc && !reconstructChroma), cancellationToken: options.CancellationToken,
                retainedManagedBytes: retained + jpeg.Length + chromaScratch, preserveRaw16: sampleBytes == 2, samplesLittleEndian: littleEndian) || jw != decodeWidth || jh != decodeRows || jc != channels || decoded.Length != expected) return false;
            if (!retainPixels) continue;
            if (reconstructChroma) ReconstructTiffJpegChroma(decoded, decodeWidth, columns, rows, horizontal, vertical,
                positioning, samples, sampleBytes, littleEndian, chromaFractions!, top * width + left, width, options);
            if (planar == 2) {
                if (ycc && (plane == 1 || plane == 2)) CopyTiffJpegChroma(decoded, source, plane, width, left, top, columns, rows,
                    decodeWidth, decodeRows, horizontal, vertical, positioning, samples, sampleBytes, littleEndian,
                    chromaFractions, top * width + left, width, options);
                else if (strips) CopyPlanarRows(decoded, source, plane, samples, sampleBytes, width, top, rows, options);
                else CopyTile(decoded, source, plane, planar, samples, sampleBytes, width, height, left, top, sw, sh, options);
            } else for (int row = 0; row < rows; row++)
                CopyWithCancellation(decoded, row * sw * samples * sampleBytes, source, ((top + row) * width + left) * samples * sampleBytes,
                    columns * samples * sampleBytes, options.CancellationToken);
        }
        return true;
    }

    private static bool TryReadJpegRationals(byte[] bytes, IReadOnlyDictionary<int, TiffEntry> entries,
        int tag, bool littleEndian, double[] values, bool allowInteger = false) {
        if (!entries.TryGetValue(tag, out TiffEntry entry)) return true;
        if (allowInteger && (entry.Type == 3 || entry.Type == 4)) {
            if (!TryReadValues(bytes, entries, tag, littleEndian, values.Length, out int[] integers)) return false;
            for (int i = 0; i < values.Length; i++) values[i] = integers[i];
            return true;
        }
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

}

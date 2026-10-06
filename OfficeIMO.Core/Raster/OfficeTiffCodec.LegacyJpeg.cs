using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    // At most four component-specific 8-bit quantization and Huffman table pairs,
    // plus frame/scan/restart headers. Reserved before any reconstructed allocation.
    private const int LegacyJpegHeaderLimit = 4096;

    private static bool TryGetLegacyJpegInterchange(byte[] bytes,
        IReadOnlyDictionary<int, TiffEntry> entries, bool littleEndian, out int offset, out int length) {
        offset = length = 0;
        if (!entries.ContainsKey(513) && !entries.ContainsKey(514)) return true;
        if (!TryReadScalar(bytes, entries, 513, littleEndian, out int start) ||
            !TryReadScalar(bytes, entries, 514, littleEndian, out int count) ||
            !HasSegment(bytes, start, count) || count < 4) return false;
        // A complete interchange image can stand alone. A partial header leaves
        // the striles and component-specific TIFF table tags authoritative.
        if (bytes[start] == 255 && bytes[start + 1] == 216 &&
            bytes[start + count - 2] == 255 && bytes[start + count - 1] == 217) {
            offset = start;
            length = count;
        }
        return true;
    }

    private static bool TryReconstructLegacyJpeg(byte[] bytes, int offset, int length,
        IReadOnlyDictionary<int, TiffEntry> entries, bool littleEndian,
        int width, int height, int samples, int channels, int plane, int precision,
        int process, int horizontal, int vertical, OfficeRasterDecodeOptions options, out byte[] jpeg) {
        jpeg = Array.Empty<byte>();
        options.CancellationToken.ThrowIfCancellationRequested();
        if (!HasSegment(bytes, offset, length) || length < 1) return false;
        if (length >= 2 && bytes[offset] == 255 && bytes[offset + 1] == 216) {
            jpeg = new byte[length];
            CopyWithCancellation(bytes, offset, jpeg, 0, length, options.CancellationToken);
            return true;
        }
        if (samples < 1 || samples > 4 || channels < 1 || channels > 4 || width > ushort.MaxValue || height > ushort.MaxValue ||
            !TryReadScalarOrDefault(bytes, entries, 515, littleEndian, 0, out int restart) || restart > ushort.MaxValue ||
            !TryReadValues(bytes, entries, 520, littleEndian, samples, out int[] dc)) return false;
        bool lossless = process == 14;
        int[] quant = Array.Empty<int>(), ac = Array.Empty<int>();
        int predictor = 0, point = 0;
        if (lossless) {
            if (!TryReadValues(bytes, entries, 517, littleEndian, samples, out int[] predictors) ||
                !TryReadValues(bytes, entries, 518, littleEndian, samples, out int[] points)) return false;
            predictor = predictors[plane < 0 ? 0 : plane];
            point = points[plane < 0 ? 0 : plane];
            if (predictor < 1 || predictor > 7 || point < 0 || point >= precision) return false;
            if (plane < 0 && (Array.Exists(predictors, p => p != predictor) || Array.Exists(points, p => p != point))) return false;
        } else if (!TryReadValues(bytes, entries, 519, littleEndian, samples, out quant) ||
            !TryReadValues(bytes, entries, 521, littleEndian, samples, out ac)) return false;

        // Some writers retain an SOS at the first strile. The per-component TIFF
        // table pointers define its tables; reconstruct canonical table/scan IDs.
        if (length >= 2 && bytes[offset] == 255 && bytes[offset + 1] == 218) {
            if (length < 4) return false;
            int scanLength = bytes[offset + 2] * 256 + bytes[offset + 3];
            if (scanLength != 6 + channels * 2 || scanLength + 2 > length || bytes[offset + 4] != channels) return false;
            int parameters = offset + 5 + channels * 2;
            if (bytes[parameters] != predictor || bytes[parameters + 1] != (lossless ? 0 : 63) || bytes[parameters + 2] != point) return false;
            for (int c = 0; c < channels; c++) {
                int at = offset + 5 + c * 2;
                if ((bytes[at + 1] >> 4) > 3 || (bytes[at + 1] & 15) > 3) return false;
                for (int prior = 0; prior < c; prior++) if (bytes[offset + 5 + prior * 2] == bytes[at]) return false;
            }
            offset += scanLength + 2;
            length -= scanLength + 2;
        }
        if (length >= 2 && bytes[offset + length - 2] == 255 && bytes[offset + length - 1] == 217) length -= 2;
        if (length < 1) return false;

        var header = new List<byte>(LegacyJpegHeaderLimit) { 255, 216 };
        for (int c = 0; c < channels; c++) {
            int source = plane < 0 ? c : plane;
            if (!lossless) {
                if (!HasBytes(bytes, quant[source], 64)) return false;
                AppendLegacyJpegMarker(header, 219, 67);
                header.Add((byte)c);
                for (int i = 0; i < 64; i++) header.Add(bytes[quant[source] + i]);
            }
            if (!TryAppendLegacyJpegHuffman(header, bytes, dc[source], c)) return false;
            if (!lossless && !TryAppendLegacyJpegHuffman(header, bytes, ac[source], 16 + c)) return false;
        }
        AppendLegacyJpegMarker(header, lossless ? 195 : precision == 12 ? 193 : 192, 8 + channels * 3);
        header.Add((byte)precision);
        AppendLegacyJpegWord(header, height);
        AppendLegacyJpegWord(header, width);
        header.Add((byte)channels);
        for (int c = 0; c < channels; c++) {
            header.Add((byte)(c + 1));
            header.Add((byte)(c == 0 || c >= 3 ? horizontal * 16 + vertical : 17));
            header.Add((byte)(lossless ? 0 : c));
        }
        if (restart != 0) {
            AppendLegacyJpegMarker(header, 221, 4);
            AppendLegacyJpegWord(header, restart);
        }
        AppendLegacyJpegMarker(header, 218, 6 + channels * 2);
        header.Add((byte)channels);
        for (int c = 0; c < channels; c++) {
            header.Add((byte)(c + 1));
            header.Add((byte)(c * 16 + (lossless ? 0 : c)));
        }
        header.Add((byte)predictor);
        header.Add((byte)(lossless ? 0 : 63));
        header.Add((byte)point);
        if (header.Count + 2 > LegacyJpegHeaderLimit) return false;
        jpeg = new byte[OfficeRasterGuards.EnsureByteCount((long)header.Count + length + 2,
            "Reconstructed TIFF JPEG exceeds the managed limit.")];
        header.CopyTo(jpeg);
        CopyWithCancellation(bytes, offset, jpeg, header.Count, length, options.CancellationToken);
        jpeg[jpeg.Length - 2] = 255;
        jpeg[jpeg.Length - 1] = 217;
        return true;
    }

    private static bool TryAppendLegacyJpegHuffman(List<byte> header, byte[] bytes, int offset, int selector) {
        if (!HasBytes(bytes, offset, 16)) return false;
        int symbols = 0;
        for (int i = 0; i < 16; i++) symbols += bytes[offset + i];
        if (symbols < 1 || symbols > 256 || !HasBytes(bytes, offset, 16 + symbols)) return false;
        AppendLegacyJpegMarker(header, 196, 19 + symbols);
        header.Add((byte)selector);
        for (int i = 0; i < 16 + symbols; i++) header.Add(bytes[offset + i]);
        return true;
    }

    private static void AppendLegacyJpegMarker(List<byte> header, int marker, int length) {
        header.Add(255);
        header.Add((byte)marker);
        AppendLegacyJpegWord(header, length);
    }

    private static void AppendLegacyJpegWord(List<byte> header, int value) {
        header.Add((byte)(value >> 8));
        header.Add((byte)value);
    }
}

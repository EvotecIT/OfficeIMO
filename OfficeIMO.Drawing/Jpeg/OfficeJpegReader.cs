using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Decodes baseline and progressive JPEG images to RGBA buffers (SOF0/SOF2, 8-bit, Huffman).
/// </summary>
internal static partial class OfficeJpegReader {
    private static readonly byte[] ZigZag = {
        0, 1, 8, 16, 9, 2, 3, 10,
        17, 24, 32, 25, 18, 11, 4, 5,
        12, 19, 26, 33, 40, 48, 41, 34,
        27, 20, 13, 6, 7, 14, 21, 28,
        35, 42, 49, 56, 57, 50, 43, 36,
        29, 22, 15, 23, 30, 37, 44, 51,
        58, 59, 52, 45, 38, 31, 39, 46,
        53, 60, 61, 54, 47, 55, 62, 63,
    };

    private static readonly byte[] StdDcLumaBits = { 0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0 };
    private static readonly byte[] StdDcChromaBits = { 0, 3, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0 };
    private static readonly byte[] StdDcValues = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 };

    private static readonly byte[] StdAcLumaBits = { 0, 2, 1, 3, 3, 2, 4, 3, 5, 5, 4, 4, 0, 0, 1, 0x7d };
    private static readonly byte[] StdAcChromaBits = { 0, 2, 1, 2, 4, 4, 3, 4, 7, 5, 4, 4, 0, 1, 2, 0x77 };

    private static readonly byte[] StdAcLumaValues = {
        0x01, 0x02, 0x03, 0x00, 0x04, 0x11, 0x05, 0x12, 0x21, 0x31, 0x41, 0x06, 0x13, 0x51, 0x61, 0x07,
        0x22, 0x71, 0x14, 0x32, 0x81, 0x91, 0xA1, 0x08, 0x23, 0x42, 0xB1, 0xC1, 0x15, 0x52, 0xD1, 0xF0,
        0x24, 0x33, 0x62, 0x72, 0x82, 0x09, 0x0A, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x25, 0x26, 0x27, 0x28,
        0x29, 0x2A, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49,
        0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69,
        0x6A, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89,
        0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9A, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7,
        0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xC2, 0xC3, 0xC4, 0xC5,
        0xC6, 0xC7, 0xC8, 0xC9, 0xCA, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xE1, 0xE2,
        0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8, 0xE9, 0xEA, 0xF1, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8,
        0xF9, 0xFA,
    };

    private static readonly byte[] StdAcChromaValues = {
        0x00, 0x01, 0x02, 0x03, 0x11, 0x04, 0x05, 0x21, 0x31, 0x06, 0x12, 0x41, 0x51, 0x07, 0x61, 0x71,
        0x13, 0x22, 0x32, 0x81, 0x08, 0x14, 0x42, 0x91, 0xA1, 0xB1, 0xC1, 0x09, 0x23, 0x33, 0x52, 0xF0,
        0x15, 0x62, 0x72, 0xD1, 0x0A, 0x16, 0x24, 0x34, 0xE1, 0x25, 0xF1, 0x17, 0x18, 0x19, 0x1A, 0x26,
        0x27, 0x28, 0x29, 0x2A, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48,
        0x49, 0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59, 0x5A, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68,
        0x69, 0x6A, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x82, 0x83, 0x84, 0x85, 0x86, 0x87,
        0x88, 0x89, 0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9A, 0xA2, 0xA3, 0xA4, 0xA5,
        0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xC2, 0xC3,
        0xC4, 0xC5, 0xC6, 0xC7, 0xC8, 0xC9, 0xCA, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA,
        0xE2, 0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8, 0xE9, 0xEA, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8,
        0xF9, 0xFA,
    };

    private static readonly double[,] IdctCos = BuildCosTable();

    private static void EnsureStandardHuffmanTables(HuffmanTable[] dcTables, HuffmanTable[] acTables) {
        if (dcTables.Length > 0 && !dcTables[0].IsValid) {
            dcTables[0] = HuffmanTable.Build(StdDcLumaBits, StdDcValues);
        }
        if (dcTables.Length > 1 && !dcTables[1].IsValid) {
            dcTables[1] = HuffmanTable.Build(StdDcChromaBits, StdDcValues);
        }
        if (acTables.Length > 0 && !acTables[0].IsValid) {
            acTables[0] = HuffmanTable.Build(StdAcLumaBits, StdAcLumaValues);
        }
        if (acTables.Length > 1 && !acTables[1].IsValid) {
            acTables[1] = HuffmanTable.Build(StdAcChromaBits, StdAcChromaValues);
        }
    }

    /// <summary>
    /// Returns true when the buffer looks like a JPEG.
    /// </summary>
    public static bool IsJpeg(byte[] data) {
        return data.Length >= 2 && data[0] == 0xFF && data[1] == 0xD8;
    }

    /// <summary>
    /// Decodes a JPEG image to an RGBA buffer.
    /// </summary>
    public static byte[] DecodeRgba32(byte[] data, out int width, out int height) {
        return DecodeRgba32(data, out width, out height, default);
    }

    /// <summary>
    /// Decodes a JPEG image to an RGBA buffer.
    /// </summary>
    public static byte[] DecodeRgba32(byte[] data, out int width, out int height, OfficeJpegDecodeOptions options) {
        if (!IsJpeg(data)) throw new FormatException("Invalid JPEG signature.");
        OfficeRasterGuards.EnsurePayloadWithinLimits(data.Length, "JPEG payload exceeds size limits.");

        var quantTables = new int[4][];
        var dcTables = new HuffmanTable[4];
        var acTables = new HuffmanTable[4];
        var restartInterval = 0;
        var hasFrame = false;
        var progressive = false;
        var orientation = 1;
        int? adobeTransform = null;
        var frame = default(JpegFrame);
        BaselineState? baselineState = null;
        ProgressiveState? progressiveState = null;

        var offset = 2;
        while (offset < data.Length) {
            if (data[offset] != 0xFF) {
                offset++;
                continue;
            }

            while (offset < data.Length && data[offset] == 0xFF) offset++;
            if (offset >= data.Length) break;
            var marker = data[offset++];

            if (marker == 0xD9) break;

            if (marker == 0xDA) {
                if (!hasFrame) throw new FormatException("Missing JPEG frame segment.");
                var segLen = ReadUInt16BE(data, offset);
                offset += 2;
                if (segLen < 2 || offset + segLen - 2 > data.Length) throw new FormatException("Invalid JPEG scan header.");
                var scan = ParseScanHeader(data.Slice(offset, segLen - 2), ref frame);
                offset += segLen - 2;

                var scanEnd = FindScanEnd(data, offset);
                var scanData = data.Slice(offset, scanEnd - offset);

                if (!progressive) {
                    ValidateBaselineScan(scan, frame, quantTables, dcTables, acTables);
                    baselineState ??= BaselineState.Create(frame);
                    DecodeBaselineScan(
                        scanData,
                        scan,
                        frame,
                        baselineState,
                        quantTables,
                        dcTables,
                        acTables,
                        restartInterval,
                        options.AllowTruncated);
                    offset = scanEnd;
                    continue;
                }

                ValidateProgressiveScan(scan, frame, quantTables, dcTables, acTables);
                progressiveState ??= ProgressiveState.Create(frame, quantTables);
                DecodeProgressiveScan(
                    scanData,
                    scan,
                    frame,
                    progressiveState,
                    quantTables,
                    dcTables,
                    acTables,
                    restartInterval,
                    options.AllowTruncated);
                offset = scanEnd;
                continue;
            }

            if (marker == 0xDB) {
                var segLen = ReadUInt16BE(data, offset);
                offset += 2;
                var end = offset + segLen - 2;
                if (segLen < 2 || end > data.Length) throw new FormatException("Invalid JPEG DQT segment.");
                while (offset < end) {
                    var info = data[offset++];
                    var precision = info >> 4;
                    var tableId = info & 0x0F;
                    if (precision != 0) throw new FormatException("Unsupported JPEG quantization precision.");
                    if (tableId >= quantTables.Length) throw new FormatException("Unsupported JPEG quantization table.");
                    if (offset + 64 > end) throw new FormatException("Invalid JPEG quantization table.");
                    var table = new int[64];
                    for (var i = 0; i < 64; i++) {
                        table[ZigZag[i]] = data[offset++];
                    }
                    quantTables[tableId] = table;
                }
                continue;
            }

            if (marker == 0xC4) {
                var segLen = ReadUInt16BE(data, offset);
                offset += 2;
                var end = offset + segLen - 2;
                if (segLen < 2 || end > data.Length) throw new FormatException("Invalid JPEG DHT segment.");
                while (offset < end) {
                    var info = data[offset++];
                    var tableClass = info >> 4;
                    var tableId = info & 0x0F;
                    if (tableId >= 4) throw new FormatException("Unsupported JPEG Huffman table.");
                    if (offset + 16 > end) throw new FormatException("Invalid JPEG Huffman table.");
                    var counts = new byte[16];
                    for (var i = 0; i < 16; i++) counts[i] = data[offset++];
                    var total = 0;
                    for (var i = 0; i < 16; i++) total += counts[i];
                    if (offset + total > end) throw new FormatException("Invalid JPEG Huffman values.");
                    var values = data.Slice(offset, total).ToArray();
                    offset += total;
                    var table = HuffmanTable.Build(counts, values);
                    if (tableClass == 0) dcTables[tableId] = table;
                    else if (tableClass == 1) acTables[tableId] = table;
                    else throw new FormatException("Unsupported JPEG Huffman table class.");
                }
                continue;
            }

            if (marker == 0xC0 || marker == 0xC2) {
                var segLen = ReadUInt16BE(data, offset);
                offset += 2;
                if (segLen < 8 || offset + segLen - 2 > data.Length) throw new FormatException("Invalid JPEG SOF segment.");
                frame = ParseFrameHeader(data.Slice(offset, segLen - 2));
                hasFrame = true;
                progressive = marker == 0xC2;
                offset += segLen - 2;
                continue;
            }

            if (marker == 0xDD) {
                var segLen = ReadUInt16BE(data, offset);
                offset += 2;
                if (segLen != 4 || offset + 2 > data.Length) throw new FormatException("Invalid JPEG DRI segment.");
                restartInterval = ReadUInt16BE(data, offset);
                offset += 2;
                continue;
            }

            if (marker == 0xE1) {
                var segLen = ReadUInt16BE(data, offset);
                offset += 2;
                if (segLen < 2 || offset + segLen - 2 > data.Length) throw new FormatException("Invalid JPEG APP1 segment.");
                var app1 = data.Slice(offset, segLen - 2);
                if (TryReadExifOrientation(app1, out var exifOrientation)) orientation = exifOrientation;
                offset += segLen - 2;
                continue;
            }

            if (marker == 0xEE) {
                var segLen = ReadUInt16BE(data, offset);
                offset += 2;
                if (segLen < 2 || offset + segLen - 2 > data.Length) throw new FormatException("Invalid JPEG APP14 segment.");
                var app14 = data.Slice(offset, segLen - 2);
                if (TryReadAdobeTransform(app14, out var transform)) adobeTransform = transform;
                offset += segLen - 2;
                continue;
            }

            if (marker >= 0xD0 && marker <= 0xD7) {
                continue;
            }

            var length = ReadUInt16BE(data, offset);
            offset += 2;
            if (length < 2 || offset + length - 2 > data.Length) throw new FormatException("Invalid JPEG segment.");
            offset += length - 2;
        }

        if (!progressive && hasFrame && baselineState is not null) {
            width = frame.Width;
            height = frame.Height;
            var rgba = baselineState.RenderRgba(frame, adobeTransform, options.HighQualityChroma);
            return ApplyOrientation(rgba, ref width, ref height, orientation);
        }

        if (progressive && hasFrame && progressiveState is not null) {
            width = frame.Width;
            height = frame.Height;
            var rgba = progressiveState.RenderRgba(frame, adobeTransform, options.HighQualityChroma);
            return ApplyOrientation(rgba, ref width, ref height, orientation);
        }

        throw new FormatException("JPEG scan not found.");
    }
}

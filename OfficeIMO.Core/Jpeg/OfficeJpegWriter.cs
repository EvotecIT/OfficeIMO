using System;
using System.IO;

namespace OfficeIMO.Drawing;

internal static class OfficeJpegWriter {
    private const string JpegOutputLimitMessage = "JPEG output exceeds size limits.";
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

    private static readonly byte[] StdLumaQuant = {
        16, 11, 10, 16, 24, 40, 51, 61,
        12, 12, 14, 19, 26, 58, 60, 55,
        14, 13, 16, 24, 40, 57, 69, 56,
        14, 17, 22, 29, 51, 87, 80, 62,
        18, 22, 37, 56, 68, 109, 103, 77,
        24, 35, 55, 64, 81, 104, 113, 92,
        49, 64, 78, 87, 103, 121, 120, 101,
        72, 92, 95, 98, 112, 100, 103, 99,
    };

    private static readonly byte[] StdChromaQuant = {
        17, 18, 24, 47, 99, 99, 99, 99,
        18, 21, 26, 66, 99, 99, 99, 99,
        24, 26, 56, 99, 99, 99, 99, 99,
        47, 66, 99, 99, 99, 99, 99, 99,
        99, 99, 99, 99, 99, 99, 99, 99,
        99, 99, 99, 99, 99, 99, 99, 99,
        99, 99, 99, 99, 99, 99, 99, 99,
        99, 99, 99, 99, 99, 99, 99, 99,
    };

    private static readonly byte[] DcLumaBits = { 0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0 };
    private static readonly byte[] DcChromaBits = { 0, 3, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0 };
    private static readonly byte[] DcValues = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 };

    private static readonly byte[] AcLumaBits = { 0, 2, 1, 3, 3, 2, 4, 3, 5, 5, 4, 4, 0, 0, 1, 0x7d };
    private static readonly byte[] AcChromaBits = { 0, 2, 1, 2, 4, 4, 3, 4, 7, 5, 4, 4, 0, 1, 2, 0x77 };

    private static readonly byte[] AcLumaValues = {
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

    private static readonly byte[] AcChromaValues = {
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

    private static readonly double[,] CosTable = BuildCosTable();
    private static readonly HuffmanTable DcLumaTable = BuildHuffmanTable(DcLumaBits, DcValues);
    private static readonly HuffmanTable AcLumaTable = BuildHuffmanTable(AcLumaBits, AcLumaValues);
    private static readonly HuffmanTable DcChromaTable = BuildHuffmanTable(DcChromaBits, DcValues);
    private static readonly HuffmanTable AcChromaTable = BuildHuffmanTable(AcChromaBits, AcChromaValues);

    public static byte[] WriteRgba(int width, int height, byte[] rgba, int stride, int quality) {
        using var ms = new MemoryStream();
        WriteRgba(ms, width, height, rgba, stride, quality);
        return ms.ToArray();
    }

    public static byte[] WriteRgbaScanlines(int width, int height, byte[] scanlines, int stride, int quality) {
        using var ms = new MemoryStream();
        WriteRgbaScanlines(ms, width, height, scanlines, stride, quality);
        return ms.ToArray();
    }

    public static byte[] WriteRgba(int width, int height, byte[] rgba, int stride, OfficeJpegEncodeOptions options) {
        using var ms = new MemoryStream();
        WriteRgba(ms, width, height, rgba, stride, options);
        return ms.ToArray();
    }

    public static byte[] WriteRgbaScanlines(int width, int height, byte[] scanlines, int stride, OfficeJpegEncodeOptions options) {
        using var ms = new MemoryStream();
        WriteRgbaScanlines(ms, width, height, scanlines, stride, options);
        return ms.ToArray();
    }

    public static void WriteRgba(Stream stream, int width, int height, byte[] rgba, int stride, int quality) {
        WriteRgbaCore(stream, width, height, rgba, stride, rowOffset: 0, rowStride: stride, BuildOptions(quality), nameof(rgba), "RGBA buffer too small.");
    }

    public static void WriteRgbaScanlines(Stream stream, int width, int height, byte[] scanlines, int stride, int quality) {
        WriteRgbaCore(stream, width, height, scanlines, stride, rowOffset: 1, rowStride: stride + 1, BuildOptions(quality), nameof(scanlines), "Scanline buffer too small.");
    }

    public static void WriteRgba(Stream stream, int width, int height, byte[] rgba, int stride, OfficeJpegEncodeOptions options) {
        WriteRgbaCore(stream, width, height, rgba, stride, rowOffset: 0, rowStride: stride, options, nameof(rgba), "RGBA buffer too small.");
    }

    public static void WriteRgbaScanlines(Stream stream, int width, int height, byte[] scanlines, int stride, OfficeJpegEncodeOptions options) {
        WriteRgbaCore(stream, width, height, scanlines, stride, rowOffset: 1, rowStride: stride + 1, options, nameof(scanlines), "Scanline buffer too small.");
    }

    private static void WriteRgbaCore(Stream stream, int width, int height, byte[] rgba, int stride, int rowOffset, int rowStride, OfficeJpegEncodeOptions options, string bufferName, string bufferMessage) {
        if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        if (width > OfficeRasterImageEncoder.JpegMaximumDimension) throw new ArgumentOutOfRangeException(nameof(width), "JPEG width cannot exceed 65535 pixels.");
        if (height > OfficeRasterImageEncoder.JpegMaximumDimension) throw new ArgumentOutOfRangeException(nameof(height), "JPEG height cannot exceed 65535 pixels.");
        _ = OfficeRasterGuards.EnsureOutputPixels(width, height, JpegOutputLimitMessage);
        _ = OfficeRasterGuards.EnsureOutputBytes((long)width * height * 4, JpegOutputLimitMessage);
        if (rgba is null) throw new ArgumentNullException(bufferName);
        if (stride < width * 4) throw new ArgumentOutOfRangeException(nameof(stride));
        if (rowStride < rowOffset + stride) throw new ArgumentOutOfRangeException(nameof(rowStride));
        if (rgba.Length < height * rowStride) throw new ArgumentException(bufferMessage, bufferName);
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (options is null) throw new ArgumentNullException(nameof(options));

        var quality = options.Quality;
        if (quality is < 1 or > 100) throw new ArgumentOutOfRangeException(nameof(options.Quality));

        var qY = ScaleQuantTable(StdLumaQuant, quality);
        var qC = ScaleQuantTable(StdChromaQuant, quality);

        var grayscale = IsGrayscale(rgba, width, height, stride, rowOffset, rowStride);
        var sampling = grayscale ? OfficeJpegSubsampling.Y444 : options.Subsampling;

        BuildComponents(sampling, grayscale, out var components, out var maxH, out var maxV);
        var coeffs = BuildCoefficients(
            rgba,
            width,
            height,
            stride,
            rowOffset,
            rowStride,
            components,
            maxH,
            maxV,
            qY,
            qC);

        var tables = BuildHuffmanTables(coeffs, components, options.OptimizeHuffman);

        WriteMarker(stream, 0xFFD8);
        if (options.WriteJfifHeader) {
            WriteApp0(stream, options.DpiX, options.DpiY);
        }
        WriteMetadata(stream, options.Metadata);
        WriteDqt(stream, 0, qY);
        if (!grayscale) {
            WriteDqt(stream, 1, qC);
        }

        WriteSof(stream, options.Progressive, width, height, components);

        WriteDht(stream, 0, 0, tables.DcLuma.Bits, tables.DcLuma.Values);
        WriteDht(stream, 1, 0, tables.AcLuma.Bits, tables.AcLuma.Values);
        if (!grayscale) {
            WriteDht(stream, 0, 1, tables.DcChroma.Bits, tables.DcChroma.Values);
            WriteDht(stream, 1, 1, tables.AcChroma.Bits, tables.AcChroma.Values);
        }

        if (options.Progressive) {
            EncodeProgressive(stream, width, height, maxH, maxV, components, coeffs, tables);
        } else {
            EncodeBaseline(stream, components, coeffs, tables);
        }

        WriteMarker(stream, 0xFFD9);
    }

    private static OfficeJpegEncodeOptions BuildOptions(int quality) {
        return new OfficeJpegEncodeOptions {
            Quality = quality
        };
    }

    private sealed class ComponentSpec {
        public byte Id;
        public byte H;
        public byte V;
        public byte QuantId;
        public byte DcTable;
        public byte AcTable;
    }

    private readonly struct ComponentCoefficients {
        public readonly ComponentSpec Spec;
        public readonly int BlocksPerRow;
        public readonly int BlocksPerCol;
        public readonly int[] Data;

        public ComponentCoefficients(ComponentSpec spec, int blocksPerRow, int blocksPerCol, int[] data) {
            Spec = spec;
            BlocksPerRow = blocksPerRow;
            BlocksPerCol = blocksPerCol;
            Data = data;
        }
    }

    private readonly struct HuffmanSpec {
        public readonly byte[] Bits;
        public readonly byte[] Values;
        public readonly HuffmanTable Table;

        public HuffmanSpec(byte[] bits, byte[] values, HuffmanTable table) {
            Bits = bits;
            Values = values;
            Table = table;
        }
    }

    private readonly struct HuffmanTableSet {
        public readonly HuffmanSpec DcLuma;
        public readonly HuffmanSpec AcLuma;
        public readonly HuffmanSpec DcChroma;
        public readonly HuffmanSpec AcChroma;

        public HuffmanTableSet(HuffmanSpec dcLuma, HuffmanSpec acLuma, HuffmanSpec dcChroma, HuffmanSpec acChroma) {
            DcLuma = dcLuma;
            AcLuma = acLuma;
            DcChroma = dcChroma;
            AcChroma = acChroma;
        }
    }

    private static void BuildComponents(OfficeJpegSubsampling subsampling, bool grayscale, out ComponentSpec[] components, out int maxH, out int maxV) {
        if (grayscale) {
            components = new[] {
                new ComponentSpec { Id = 1, H = 1, V = 1, QuantId = 0, DcTable = 0, AcTable = 0 }
            };
            maxH = 1;
            maxV = 1;
            return;
        }

        byte yH;
        byte yV;
        byte cH;
        byte cV;
        switch (subsampling) {
            case OfficeJpegSubsampling.Y422:
                yH = 2;
                yV = 1;
                cH = 1;
                cV = 1;
                break;
            case OfficeJpegSubsampling.Y420:
                yH = 2;
                yV = 2;
                cH = 1;
                cV = 1;
                break;
            case OfficeJpegSubsampling.Y444:
            default:
                yH = 1;
                yV = 1;
                cH = 1;
                cV = 1;
                break;
        }

        maxH = yH;
        maxV = yV;
        components = new[] {
            new ComponentSpec { Id = 1, H = yH, V = yV, QuantId = 0, DcTable = 0, AcTable = 0 },
            new ComponentSpec { Id = 2, H = cH, V = cV, QuantId = 1, DcTable = 1, AcTable = 1 },
            new ComponentSpec { Id = 3, H = cH, V = cV, QuantId = 1, DcTable = 1, AcTable = 1 }
        };
    }

    private static ComponentCoefficients[] BuildCoefficients(
        byte[] rgba,
        int width,
        int height,
        int stride,
        int rowOffset,
        int rowStride,
        ComponentSpec[] components,
        int maxH,
        int maxV,
        int[] qY,
        int[] qC) {
        var mcuWidth = maxH * 8;
        var mcuHeight = maxV * 8;
        var mcuCols = (width + mcuWidth - 1) / mcuWidth;
        var mcuRows = (height + mcuHeight - 1) / mcuHeight;

        var result = new ComponentCoefficients[components.Length];
        for (var i = 0; i < components.Length; i++) {
            var comp = components[i];
            var blocksPerRow = mcuCols * comp.H;
            var blocksPerCol = mcuRows * comp.V;
            result[i] = new ComponentCoefficients(comp, blocksPerRow, blocksPerCol, new int[blocksPerRow * blocksPerCol * 64]);
        }

        var yBlock = new int[64];
        var cbBlock = new int[64];
        var crBlock = new int[64];
        var temp = new int[64];

        ComponentCoefficients? yCoeffs = null;
        ComponentCoefficients? cbCoeffs = null;
        ComponentCoefficients? crCoeffs = null;

        for (var i = 0; i < result.Length; i++) {
            var comp = result[i];
            if (comp.Spec.Id == 1) yCoeffs = comp;
            else if (comp.Spec.Id == 2) cbCoeffs = comp;
            else if (comp.Spec.Id == 3) crCoeffs = comp;
        }

        var hasChroma = cbCoeffs.HasValue && crCoeffs.HasValue;

        for (var my = 0; my < mcuRows; my++) {
            for (var mx = 0; mx < mcuCols; mx++) {
                if (yCoeffs.HasValue) {
                    var yc = yCoeffs.Value;
                    for (var vy = 0; vy < yc.Spec.V; vy++) {
                        for (var vx = 0; vx < yc.Spec.H; vx++) {
                            var blockX = mx * yc.Spec.H + vx;
                            var blockY = my * yc.Spec.V + vy;
                            var x0 = blockX * 8;
                            var y0 = blockY * 8;
                            LoadBlockLuma(rgba, stride, rowOffset, rowStride, width, height, x0, y0, yBlock);
                            ForwardDctQuantize(yBlock, qY, temp);
                            var offset = (blockY * yc.BlocksPerRow + blockX) * 64;
                            Array.Copy(temp, 0, yc.Data, offset, 64);
                        }
                    }
                }

                if (hasChroma) {
                    var cb = cbCoeffs!.Value;
                    var cr = crCoeffs!.Value;
                    var sampleW = maxH / cb.Spec.H;
                    var sampleH = maxV / cb.Spec.V;
                    var blockX = mx * cb.Spec.H;
                    var blockY = my * cb.Spec.V;
                    var x0 = blockX * 8 * sampleW;
                    var y0 = blockY * 8 * sampleH;

                    LoadBlockChroma(rgba, stride, rowOffset, rowStride, width, height, x0, y0, sampleW, sampleH, cbBlock, crBlock);

                    ForwardDctQuantize(cbBlock, qC, temp);
                    var offsetCb = (blockY * cb.BlocksPerRow + blockX) * 64;
                    Array.Copy(temp, 0, cb.Data, offsetCb, 64);

                    ForwardDctQuantize(crBlock, qC, temp);
                    var offsetCr = (blockY * cr.BlocksPerRow + blockX) * 64;
                    Array.Copy(temp, 0, cr.Data, offsetCr, 64);
                }
            }
        }

        return result;
    }

    private static HuffmanTableSet BuildHuffmanTables(ComponentCoefficients[] coeffs, ComponentSpec[] components, bool optimize) {
        if (!optimize) {
            var dcLuma = new HuffmanSpec(DcLumaBits, DcValues, DcLumaTable);
            var acLuma = new HuffmanSpec(AcLumaBits, AcLumaValues, AcLumaTable);
            var dcChroma = new HuffmanSpec(DcChromaBits, DcValues, DcChromaTable);
            var acChroma = new HuffmanSpec(AcChromaBits, AcChromaValues, AcChromaTable);
            return new HuffmanTableSet(dcLuma, acLuma, dcChroma, acChroma);
        }

        var freqDcLuma = new int[256];
        var freqAcLuma = new int[256];
        var freqDcChroma = new int[256];
        var freqAcChroma = new int[256];

        AccumulateFrequencies(coeffs, components, freqDcLuma, freqAcLuma, freqDcChroma, freqAcChroma);

        var dcL = BuildOptimizedHuffman(freqDcLuma);
        var acL = BuildOptimizedHuffman(freqAcLuma);
        var dcC = BuildOptimizedHuffman(freqDcChroma);
        var acC = BuildOptimizedHuffman(freqAcChroma);

        return new HuffmanTableSet(dcL, acL, dcC, acC);
    }

    private static void EncodeImage(
        BitWriter bw,
        int width,
        int height,
        byte[] rgba,
        int stride,
        int rowOffset,
        int rowStride,
        int[] qY,
        int[] qC,
        HuffmanTable dcY,
        HuffmanTable acY,
        HuffmanTable dcC,
        HuffmanTable acC) {
        var blockY = new int[64];
        var blockCb = new int[64];
        var blockCr = new int[64];
        var temp = new int[64];

        var prevY = 0;
        var prevCb = 0;
        var prevCr = 0;

        for (var by = 0; by < height; by += 8) {
            for (var bx = 0; bx < width; bx += 8) {
                LoadBlock(rgba, stride, rowOffset, rowStride, width, height, bx, by, blockY, blockCb, blockCr);

                EncodeBlock(bw, blockY, qY, dcY, acY, ref prevY, temp);
                EncodeBlock(bw, blockCb, qC, dcC, acC, ref prevCb, temp);
                EncodeBlock(bw, blockCr, qC, dcC, acC, ref prevCr, temp);
            }
        }
    }

    private static void EncodeImageGray(
        BitWriter bw,
        int width,
        int height,
        byte[] rgba,
        int stride,
        int rowOffset,
        int rowStride,
        int[] qY,
        HuffmanTable dcY,
        HuffmanTable acY) {
        var blockY = new int[64];
        var temp = new int[64];
        var prevY = 0;

        for (var by = 0; by < height; by += 8) {
            for (var bx = 0; bx < width; bx += 8) {
                LoadBlockLuma(rgba, stride, rowOffset, rowStride, width, height, bx, by, blockY);
                EncodeBlock(bw, blockY, qY, dcY, acY, ref prevY, temp);
            }
        }
    }

    private static void LoadBlock(
        byte[] rgba,
        int stride,
        int rowOffset,
        int rowStride,
        int width,
        int height,
        int bx,
        int by,
        int[] yBlock,
        int[] cbBlock,
        int[] crBlock) {
        var i = 0;
        for (var y = 0; y < 8; y++) {
            var py = by + y;
            if (py >= height) py = height - 1;
            var row = py * rowStride + rowOffset;
            for (var x = 0; x < 8; x++) {
                var px = bx + x;
                if (px >= width) px = width - 1;
                var p = row + px * 4;
                var r = rgba[p + 0];
                var g = rgba[p + 1];
                var b = rgba[p + 2];
                var a = rgba[p + 3];
                if (a != 255) {
                    var inv = 255 - a;
                    r = (byte)((r * a + 255 * inv + 127) / 255);
                    g = (byte)((g * a + 255 * inv + 127) / 255);
                    b = (byte)((b * a + 255 * inv + 127) / 255);
                }

                var yv = (77 * r + 150 * g + 29 * b + 128) >> 8;
                var cb = ((-43 * r - 85 * g + 128 * b + 128) >> 8) + 128;
                var cr = ((128 * r - 107 * g - 21 * b + 128) >> 8) + 128;

                yBlock[i] = yv - 128;
                cbBlock[i] = cb - 128;
                crBlock[i] = cr - 128;
                i++;
            }
        }
    }

    private static void LoadBlockLuma(
        byte[] rgba,
        int stride,
        int rowOffset,
        int rowStride,
        int width,
        int height,
        int bx,
        int by,
        int[] yBlock) {
        var i = 0;
        for (var y = 0; y < 8; y++) {
            var py = by + y;
            if (py >= height) py = height - 1;
            var row = py * rowStride + rowOffset;
            for (var x = 0; x < 8; x++) {
                var px = bx + x;
                if (px >= width) px = width - 1;
                var p = row + px * 4;
                var r = rgba[p + 0];
                var g = rgba[p + 1];
                var b = rgba[p + 2];
                var a = rgba[p + 3];
                if (a != 255) {
                    var inv = 255 - a;
                    r = (byte)((r * a + 255 * inv + 127) / 255);
                    g = (byte)((g * a + 255 * inv + 127) / 255);
                    b = (byte)((b * a + 255 * inv + 127) / 255);
                }

                var yv = (77 * r + 150 * g + 29 * b + 128) >> 8;
                yBlock[i] = yv - 128;
                i++;
            }
        }
    }

    private static void LoadBlockChroma(
        byte[] rgba,
        int stride,
        int rowOffset,
        int rowStride,
        int width,
        int height,
        int bx,
        int by,
        int sampleW,
        int sampleH,
        int[] cbBlock,
        int[] crBlock) {
        var i = 0;
        var count = sampleW * sampleH;
        for (var y = 0; y < 8; y++) {
            var baseY = by + y * sampleH;
            for (var x = 0; x < 8; x++) {
                var baseX = bx + x * sampleW;
                var sumR = 0;
                var sumG = 0;
                var sumB = 0;

                for (var sy = 0; sy < sampleH; sy++) {
                    var py = baseY + sy;
                    if (py >= height) py = height - 1;
                    var row = py * rowStride + rowOffset;
                    for (var sx = 0; sx < sampleW; sx++) {
                        var px = baseX + sx;
                        if (px >= width) px = width - 1;
                        var p = row + px * 4;
                        var r = rgba[p + 0];
                        var g = rgba[p + 1];
                        var b = rgba[p + 2];
                        var a = rgba[p + 3];
                        if (a != 255) {
                            var inv = 255 - a;
                            r = (byte)((r * a + 255 * inv + 127) / 255);
                            g = (byte)((g * a + 255 * inv + 127) / 255);
                            b = (byte)((b * a + 255 * inv + 127) / 255);
                        }
                        sumR += r;
                        sumG += g;
                        sumB += b;
                    }
                }

                var rAvg = (sumR + count / 2) / count;
                var gAvg = (sumG + count / 2) / count;
                var bAvg = (sumB + count / 2) / count;

                var cb = ((-43 * rAvg - 85 * gAvg + 128 * bAvg + 128) >> 8) + 128;
                var cr = ((128 * rAvg - 107 * gAvg - 21 * bAvg + 128) >> 8) + 128;

                cbBlock[i] = cb - 128;
                crBlock[i] = cr - 128;
                i++;
            }
        }
    }

    private static void EncodeBlock(
        BitWriter bw,
        int[] input,
        int[] quant,
        HuffmanTable dcTable,
        HuffmanTable acTable,
        ref int prevDc,
        int[] temp) {
        ForwardDctQuantize(input, quant, temp);

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
        int[] coeffs,
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

    private static void EncodeDcFromQuantized(BitWriter bw, int[] coeffs, int offset, HuffmanTable dcTable, ref int prevDc) {
        var dc = coeffs[offset];
        var diff = dc - prevDc;
        prevDc = dc;
        var dcCat = BitCount(diff);
        bw.WriteBits(dcTable.Codes[dcCat], dcTable.Sizes[dcCat]);
        if (dcCat > 0) {
            bw.WriteBits(EncodeValue(diff, dcCat), dcCat);
        }
    }

    private static void EncodeAcFromQuantized(BitWriter bw, int[] coeffs, int offset, HuffmanTable acTable, int ss, int se) {
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

    private static void ForwardDctQuantize(int[] input, int[] quant, int[] output) {
        const double invSqrt2 = 0.7071067811865476;
        for (var u = 0; u < 8; u++) {
            var cu = u == 0 ? invSqrt2 : 1.0;
            for (var v = 0; v < 8; v++) {
                var cv = v == 0 ? invSqrt2 : 1.0;
                double sum = 0;
                for (var x = 0; x < 8; x++) {
                    for (var y = 0; y < 8; y++) {
                        sum += input[y * 8 + x] * CosTable[u, x] * CosTable[v, y];
                    }
                }
                var coeff = 0.25 * cu * cv * sum;
                var idx = v * 8 + u;
                output[idx] = (int)Math.Round(coeff / quant[idx]);
            }
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

    private static int[] ScaleQuantTable(byte[] table, int quality) {
        var scale = quality < 50 ? 5000 / quality : 200 - quality * 2;
        var outTable = new int[64];
        for (var i = 0; i < 64; i++) {
            var val = (table[i] * scale + 50) / 100;
            if (val < 1) val = 1;
            if (val > 255) val = 255;
            outTable[i] = val;
        }
        return outTable;
    }

    private static void WriteApp0(Stream s, double dpiX, double dpiY) {
        if (dpiX < OfficeRasterImageEncoder.JpegMinimumDpi ||
            double.IsNaN(dpiX) ||
            double.IsInfinity(dpiX) ||
            dpiX > ushort.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(dpiX));
        }
        if (dpiY < OfficeRasterImageEncoder.JpegMinimumDpi ||
            double.IsNaN(dpiY) ||
            double.IsInfinity(dpiY) ||
            dpiY > ushort.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(dpiY));
        }
        WriteMarker(s, 0xFFE0);
        WriteUInt16(s, 16);
        s.WriteByte((byte)'J');
        s.WriteByte((byte)'F');
        s.WriteByte((byte)'I');
        s.WriteByte((byte)'F');
        s.WriteByte(0);
        s.WriteByte(1);
        s.WriteByte(1);
        s.WriteByte(1);
        WriteUInt16(s, checked((ushort)Math.Round(dpiX, MidpointRounding.AwayFromZero)));
        WriteUInt16(s, checked((ushort)Math.Round(dpiY, MidpointRounding.AwayFromZero)));
        s.WriteByte(0);
        s.WriteByte(0);
    }

    private static readonly byte[] ExifPrefix = { (byte)'E', (byte)'x', (byte)'i', (byte)'f', 0, 0 };
    private static readonly byte[] XmpPrefix = {
        (byte)'h', (byte)'t', (byte)'t', (byte)'p', (byte)':', (byte)'/', (byte)'/',
        (byte)'n', (byte)'s', (byte)'.', (byte)'a', (byte)'d', (byte)'o', (byte)'b', (byte)'e', (byte)'.',
        (byte)'c', (byte)'o', (byte)'m', (byte)'/', (byte)'x', (byte)'a', (byte)'p', (byte)'/',
        (byte)'1', (byte)'.', (byte)'0', (byte)'/', 0
    };
    private static readonly byte[] IccPrefix = {
        (byte)'I', (byte)'C', (byte)'C', (byte)'_', (byte)'P', (byte)'R', (byte)'O', (byte)'F', (byte)'I', (byte)'L', (byte)'E', 0
    };

    private const int MaxAppSegmentPayload = 0xFFFD;

    private static void WriteMetadata(Stream s, OfficeJpegMetadata metadata) {
        if (!metadata.HasData) return;

        byte[]? exif = metadata.Exif;
        byte[]? xmp = metadata.Xmp;
        byte[]? icc = metadata.Icc;
        if (exif is { Length: > 0 }) {
            var payload = EnsurePrefix(exif, ExifPrefix);
            WriteAppSegment(s, 0xFFE1, payload);
        }

        if (xmp is { Length: > 0 }) {
            var payload = EnsurePrefix(xmp, XmpPrefix);
            WriteAppSegment(s, 0xFFE1, payload);
        }

        if (icc is { Length: > 0 }) {
            WriteIccSegments(s, icc);
        }
    }

    private static byte[] EnsurePrefix(byte[] data, byte[] prefix) {
        if (StartsWith(data, prefix)) {
            if (data.Length > MaxAppSegmentPayload) {
                throw new ArgumentOutOfRangeException(nameof(data), "APP segment payload too large.");
            }
            return data;
        }
        if (prefix.Length + data.Length > MaxAppSegmentPayload) {
            throw new ArgumentOutOfRangeException(nameof(data), "APP segment payload too large.");
        }
        var combined = new byte[prefix.Length + data.Length];
        Buffer.BlockCopy(prefix, 0, combined, 0, prefix.Length);
        Buffer.BlockCopy(data, 0, combined, prefix.Length, data.Length);
        return combined;
    }

    private static bool StartsWith(byte[] data, byte[] prefix) {
        if (data.Length < prefix.Length) return false;
        for (var i = 0; i < prefix.Length; i++) {
            if (data[i] != prefix[i]) return false;
        }
        return true;
    }

    private static void WriteAppSegment(Stream s, int marker, byte[] payload) {
        if (payload.Length > MaxAppSegmentPayload) {
            throw new ArgumentOutOfRangeException(nameof(payload), "APP segment payload too large.");
        }
        WriteMarker(s, marker);
        WriteUInt16(s, (ushort)(payload.Length + 2));
        s.Write(payload, 0, payload.Length);
    }

    private static void WriteIccSegments(Stream s, byte[] icc) {
        const int maxPayload = MaxAppSegmentPayload;
        var headerSize = IccPrefix.Length + 2;
        var maxData = maxPayload - headerSize;
        var totalSegments = (icc.Length + maxData - 1) / maxData;
        if (totalSegments <= 0) return;
        if (totalSegments > 255) {
            throw new ArgumentOutOfRangeException(nameof(icc), "ICC profile payload too large.");
        }
        var offset = 0;
        for (var segment = 1; segment <= totalSegments; segment++) {
            var remaining = icc.Length - offset;
            var size = remaining > maxData ? maxData : remaining;
            var payload = new byte[headerSize + size];
            Buffer.BlockCopy(IccPrefix, 0, payload, 0, IccPrefix.Length);
            payload[IccPrefix.Length] = (byte)segment;
            payload[IccPrefix.Length + 1] = (byte)totalSegments;
            Buffer.BlockCopy(icc, offset, payload, headerSize, size);
            WriteAppSegment(s, 0xFFE2, payload);
            offset += size;
        }
    }

    private static void WriteDqt(Stream s, int tableId, int[] table) {
        WriteMarker(s, 0xFFDB);
        WriteUInt16(s, 67);
        s.WriteByte((byte)tableId);
        for (var i = 0; i < 64; i++) {
            s.WriteByte((byte)table[ZigZag[i]]);
        }
    }

    private static void WriteSof(Stream s, bool progressive, int width, int height, ComponentSpec[] components) {
        WriteMarker(s, progressive ? 0xFFC2 : 0xFFC0);
        var length = 8 + components.Length * 3;
        WriteUInt16(s, (ushort)length);
        s.WriteByte(8);
        WriteUInt16(s, (ushort)height);
        WriteUInt16(s, (ushort)width);
        s.WriteByte((byte)components.Length);
        for (var i = 0; i < components.Length; i++) {
            var comp = components[i];
            s.WriteByte(comp.Id);
            s.WriteByte((byte)((comp.H << 4) | comp.V));
            s.WriteByte(comp.QuantId);
        }
    }

    private static void WriteDht(Stream s, int tableClass, int tableId, byte[] bits, byte[] values) {
        WriteMarker(s, 0xFFC4);
        WriteUInt16(s, (ushort)(2 + 1 + 16 + values.Length));
        s.WriteByte((byte)((tableClass << 4) | tableId));
        for (var i = 0; i < 16; i++) s.WriteByte(bits[i]);
        s.Write(values, 0, values.Length);
    }

    private static void WriteSos(Stream s, ComponentSpec[] components, int[] componentIndices, byte ss, byte se, byte ah, byte al) {
        WriteMarker(s, 0xFFDA);
        var count = componentIndices.Length;
        WriteUInt16(s, (ushort)(6 + count * 2));
        s.WriteByte((byte)count);
        for (var i = 0; i < count; i++) {
            var comp = components[componentIndices[i]];
            s.WriteByte(comp.Id);
            s.WriteByte((byte)((comp.DcTable << 4) | comp.AcTable));
        }
        s.WriteByte(ss);
        s.WriteByte(se);
        s.WriteByte((byte)((ah << 4) | al));
    }

    private static bool IsGrayscale(byte[] rgba, int width, int height, int stride, int rowOffset, int rowStride) {
        for (var y = 0; y < height; y++) {
            var row = y * rowStride + rowOffset;
            for (var x = 0; x < width; x++) {
                var p = row + x * 4;
                var r = rgba[p + 0];
                var g = rgba[p + 1];
                var b = rgba[p + 2];
                var a = rgba[p + 3];
                if (a != 255) {
                    var inv = 255 - a;
                    r = (byte)((r * a + 255 * inv + 127) / 255);
                    g = (byte)((g * a + 255 * inv + 127) / 255);
                    b = (byte)((b * a + 255 * inv + 127) / 255);
                }
                if (r != g || r != b) return false;
            }
        }
        return true;
    }

    private static void EncodeBaseline(
        Stream stream,
        ComponentSpec[] components,
        ComponentCoefficients[] coeffs,
        HuffmanTableSet tables) {
        var mcuCols = coeffs[0].BlocksPerRow / components[0].H;
        var mcuRows = coeffs[0].BlocksPerCol / components[0].V;

        var allComponents = new int[components.Length];
        for (var i = 0; i < allComponents.Length; i++) allComponents[i] = i;
        WriteSos(stream, components, allComponents, 0, 63, 0, 0);

        var bw = new BitWriter(stream);
        var prevDc = new int[components.Length];

        for (var my = 0; my < mcuRows; my++) {
            for (var mx = 0; mx < mcuCols; mx++) {
                for (var ci = 0; ci < components.Length; ci++) {
                    var comp = components[ci];
                    var compCoeffs = coeffs[ci];
                    for (var vy = 0; vy < comp.V; vy++) {
                        for (var vx = 0; vx < comp.H; vx++) {
                            var blockX = mx * comp.H + vx;
                            var blockY = my * comp.V + vy;
                            var offset = (blockY * compCoeffs.BlocksPerRow + blockX) * 64;
                            var dcTable = ci == 0 ? tables.DcLuma.Table : tables.DcChroma.Table;
                            var acTable = ci == 0 ? tables.AcLuma.Table : tables.AcChroma.Table;
                            EncodeBlockFromQuantized(bw, compCoeffs.Data, offset, dcTable, acTable, ref prevDc[ci]);
                        }
                    }
                }
            }
        }
        bw.Flush();
    }

    private static void EncodeProgressive(
        Stream stream,
        int width,
        int height,
        int maxH,
        int maxV,
        ComponentSpec[] components,
        ComponentCoefficients[] coeffs,
        HuffmanTableSet tables) {
        var mcuCols = coeffs[0].BlocksPerRow / components[0].H;
        var mcuRows = coeffs[0].BlocksPerCol / components[0].V;

        var allComponents = new int[components.Length];
        for (var i = 0; i < allComponents.Length; i++) allComponents[i] = i;

        // DC scan
        WriteSos(stream, components, allComponents, 0, 0, 0, 0);
        var bw = new BitWriter(stream);
        var prevDc = new int[components.Length];
        for (var my = 0; my < mcuRows; my++) {
            for (var mx = 0; mx < mcuCols; mx++) {
                for (var ci = 0; ci < components.Length; ci++) {
                    var comp = components[ci];
                    var compCoeffs = coeffs[ci];
                    for (var vy = 0; vy < comp.V; vy++) {
                        for (var vx = 0; vx < comp.H; vx++) {
                            var blockX = mx * comp.H + vx;
                            var blockY = my * comp.V + vy;
                            var offset = (blockY * compCoeffs.BlocksPerRow + blockX) * 64;
                            var dcTable = ci == 0 ? tables.DcLuma.Table : tables.DcChroma.Table;
                            EncodeDcFromQuantized(bw, compCoeffs.Data, offset, dcTable, ref prevDc[ci]);
                        }
                    }
                }
            }
        }
        bw.Flush();

        // Progressive AC scans must be non-interleaved. Emit one scan per component and omit padded
        // edge blocks that belong only to interleaved MCU alignment.
        for (var ci = 0; ci < components.Length; ci++) {
            var componentIndex = new[] { ci };
            WriteSos(stream, components, componentIndex, 1, 63, 0, 0);
            bw = new BitWriter(stream);
            var comp = components[ci];
            var compCoeffs = coeffs[ci];
            var blockCols = DivideRoundUp((long)width * comp.H, maxH * 8);
            var blockRows = DivideRoundUp((long)height * comp.V, maxV * 8);
            var acTable = ci == 0 ? tables.AcLuma.Table : tables.AcChroma.Table;
            for (var blockY = 0; blockY < blockRows; blockY++) {
                for (var blockX = 0; blockX < blockCols; blockX++) {
                    var offset = (blockY * compCoeffs.BlocksPerRow + blockX) * 64;
                    EncodeAcFromQuantized(bw, compCoeffs.Data, offset, acTable, 1, 63);
                }
            }
            bw.Flush();
        }
    }

    private static int DivideRoundUp(long value, int divisor) => checked((int)((value + divisor - 1L) / divisor));

    private static void AccumulateFrequencies(
        ComponentCoefficients[] coeffs,
        ComponentSpec[] components,
        int[] dcLuma,
        int[] acLuma,
        int[] dcChroma,
        int[] acChroma) {
        var mcuCols = coeffs[0].BlocksPerRow / components[0].H;
        var mcuRows = coeffs[0].BlocksPerCol / components[0].V;
        var prevDc = new int[components.Length];

        for (var my = 0; my < mcuRows; my++) {
            for (var mx = 0; mx < mcuCols; mx++) {
                for (var ci = 0; ci < components.Length; ci++) {
                    var comp = components[ci];
                    var compCoeffs = coeffs[ci];
                    for (var vy = 0; vy < comp.V; vy++) {
                        for (var vx = 0; vx < comp.H; vx++) {
                            var blockX = mx * comp.H + vx;
                            var blockY = my * comp.V + vy;
                            var offset = (blockY * compCoeffs.BlocksPerRow + blockX) * 64;
                            var dc = compCoeffs.Data[offset];
                            var diff = dc - prevDc[ci];
                            prevDc[ci] = dc;
                            var dcCat = BitCount(diff);
                            if (ci == 0) dcLuma[dcCat]++; else dcChroma[dcCat]++;

                            AccumulateAcFrequencies(compCoeffs.Data, offset, ci == 0 ? acLuma : acChroma);
                        }
                    }
                }
            }
        }
    }

    private static void AccumulateAcFrequencies(int[] coeffs, int offset, int[] freq) {
        var zeroRun = 0;
        for (var i = 1; i < 64; i++) {
            var v = coeffs[offset + ZigZag[i]];
            if (v == 0) {
                zeroRun++;
                continue;
            }

            while (zeroRun >= 16) {
                freq[0xF0]++;
                zeroRun -= 16;
            }

            var cat = BitCount(v);
            var symbol = (zeroRun << 4) | cat;
            freq[symbol]++;
            zeroRun = 0;
        }

        if (zeroRun > 0) {
            freq[0x00]++;
        }
    }

    private static HuffmanSpec BuildOptimizedHuffman(int[] frequencies) {
        var symbols = new int[256];
        var counts = 0;
        for (var i = 0; i < frequencies.Length; i++) {
            if (frequencies[i] > 0) {
                symbols[counts++] = i;
            }
        }
        if (counts == 0) {
            symbols[counts++] = 0;
            frequencies[0] = 1;
        }

        var lengths = BuildCodeLengths(frequencies);
        var bits = new int[33];
        for (var i = 0; i < lengths.Length; i++) {
            var len = lengths[i];
            if (len > 0) bits[len]++;
        }

        LimitCodeLengths(bits, 16);

        var total = 0;
        for (var i = 1; i <= 16; i++) total += bits[i];
        if (total < counts) {
            bits[16] += counts - total;
        } else if (total > counts) {
            var extra = total - counts;
            for (var i = 16; i >= 1 && extra > 0; i--) {
                var take = Math.Min(extra, bits[i]);
                bits[i] -= take;
                extra -= take;
            }
        }

        var ordered = new int[counts];
        Array.Copy(symbols, ordered, counts);
        Array.Sort(ordered, (a, b) => {
            var fa = frequencies[a];
            var fb = frequencies[b];
            var diff = fb.CompareTo(fa);
            return diff != 0 ? diff : a.CompareTo(b);
        });

        var values = new byte[counts];
        var sizes = new byte[256];
        var index = 0;
        var valueIndex = 0;
        for (var len = 1; len <= 16; len++) {
            var count = bits[len];
            for (var i = 0; i < count && index < ordered.Length; i++) {
                var symbol = ordered[index++];
                sizes[symbol] = (byte)len;
                values[valueIndex++] = (byte)symbol;
            }
        }

        var bitsOut = new byte[16];
        for (var i = 1; i <= 16; i++) bitsOut[i - 1] = (byte)bits[i];

        var table = BuildHuffmanTable(bitsOut, values);
        return new HuffmanSpec(bitsOut, values, table);
    }

    private static int[] BuildCodeLengths(int[] frequencies) {
        var maxNodes = frequencies.Length * 2;
        var freq = new int[maxNodes];
        var parent = new int[maxNodes];
        var symbol = new int[maxNodes];
        var nodeCount = 0;
        var symbolNodes = new int[frequencies.Length];
        for (var i = 0; i < symbolNodes.Length; i++) {
            symbolNodes[i] = -1;
        }

        for (var i = 0; i < frequencies.Length; i++) {
            if (frequencies[i] <= 0) continue;
            freq[nodeCount] = frequencies[i];
            parent[nodeCount] = -1;
            symbol[nodeCount] = i;
            symbolNodes[i] = nodeCount;
            nodeCount++;
        }

        var lengthsOut = new int[frequencies.Length];
        if (nodeCount == 1) {
            lengthsOut[symbol[0]] = 1;
            return lengthsOut;
        }

        var nodesTotal = nodeCount;
        while (true) {
            var least1 = -1;
            var least2 = -1;
            for (var i = 0; i < nodesTotal; i++) {
                if (parent[i] != -1) continue;
                if (least1 < 0 || freq[i] < freq[least1]) {
                    least2 = least1;
                    least1 = i;
                } else if (least2 < 0 || freq[i] < freq[least2]) {
                    least2 = i;
                }
            }

            if (least2 < 0) break;

            freq[nodesTotal] = freq[least1] + freq[least2];
            parent[least1] = nodesTotal;
            parent[least2] = nodesTotal;
            parent[nodesTotal] = -1;
            symbol[nodesTotal] = -1;
            nodesTotal++;
        }

        for (var i = 0; i < frequencies.Length; i++) {
            var node = symbolNodes[i];
            if (node < 0) continue;
            var depth = 0;
            while (parent[node] != -1) {
                depth++;
                node = parent[node];
            }
            lengthsOut[i] = depth == 0 ? 1 : depth;
        }

        return lengthsOut;
    }

    private static void LimitCodeLengths(int[] bits, int maxLen) {
        for (var i = bits.Length - 1; i > maxLen; i--) {
            while (bits[i] > 0) {
                var j = i - 1;
                while (j > 0 && bits[j] == 0) j--;
                if (j == 0) break;
                if (bits[i] < 2) {
                    bits[i] = 0;
                    break;
                }
                bits[i] -= 2;
                bits[i - 1] += 1;
                bits[j] -= 1;
                bits[j + 1] += 2;
            }
        }
    }

    private static HuffmanTable BuildHuffmanTable(byte[] bits, byte[] values) {
        var sizes = new byte[256];
        var codes = new ushort[256];
        var code = 0;
        var k = 0;
        for (var i = 1; i <= 16; i++) {
            var count = bits[i - 1];
            for (var j = 0; j < count; j++) {
                var val = values[k++];
                sizes[val] = (byte)i;
                codes[val] = (ushort)code;
                code++;
            }
            code <<= 1;
        }
        return new HuffmanTable(codes, sizes);
    }

    private static double[,] BuildCosTable() {
        var table = new double[8, 8];
        for (var u = 0; u < 8; u++) {
            for (var x = 0; x < 8; x++) {
                table[u, x] = Math.Cos(((2 * x + 1) * u * Math.PI) / 16.0);
            }
        }
        return table;
    }

    private static void WriteMarker(Stream s, int marker) {
        s.WriteByte(0xFF);
        s.WriteByte((byte)(marker & 0xFF));
    }

    private static void WriteUInt16(Stream s, ushort value) {
        s.WriteByte((byte)(value >> 8));
        s.WriteByte((byte)(value & 0xFF));
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
            if (_bits <= 0) return;
            var b = (byte)((_buffer << (8 - _bits)) & 0xFF);
            WriteByte(b);
            _bits = 0;
        }

        private void WriteByte(byte b) {
            _stream.WriteByte(b);
            if (b == 0xFF) _stream.WriteByte(0x00);
        }
    }
}

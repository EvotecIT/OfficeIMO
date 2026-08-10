using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free GIF decoder with explicit composited-frame selection.
/// </summary>
public static class OfficeGifReader {
    /// <summary>
    /// Attempts to decode the first GIF image frame into an RGBA raster buffer.
    /// </summary>
    public static bool TryDecode(byte[]? bytes, out OfficeRasterImage? image) =>
        TryDecodeFrame(bytes, 0, out image, out _);

    /// <summary>
    /// Attempts to decode one zero-based composited GIF frame and reports the total frame count.
    /// </summary>
    public static bool TryDecodeFrame(byte[]? bytes, int frameIndex, out OfficeRasterImage? image, out int frameCount) {
        bool success = TryDecodeFrameCore(bytes, frameIndex, validateAllFrames: false, out image, out frameCount);
        if (!success) image = null;
        return success;
    }

    /// <summary>Validates every GIF frame payload while avoiding persistent frame output.</summary>
    internal static bool TryValidateAllFrames(byte[]? bytes) =>
        TryDecodeFrameCore(bytes, frameIndex: 0, validateAllFrames: true, out _, out _);

    private static bool TryDecodeFrameCore(
        byte[]? bytes,
        int frameIndex,
        bool validateAllFrames,
        out OfficeRasterImage? image,
        out int frameCount) {
        image = null;
        frameCount = 0;
        try {
            if (frameIndex < 0) return false;
            if (bytes == null || bytes.Length < 13 ||
                bytes[0] != (byte)'G' || bytes[1] != (byte)'I' || bytes[2] != (byte)'F') {
                return false;
            }
            OfficeRasterGuards.EnsurePayloadWithinLimits(bytes.Length, "GIF payload exceeds size limits.");

            string signature = GetAscii(bytes, 0, 6);
            if (signature != "GIF87a" && signature != "GIF89a") {
                return false;
            }

            int width = ReadUInt16LittleEndian(bytes, 6);
            int height = ReadUInt16LittleEndian(bytes, 8);
            if (width <= 0 || height <= 0) {
                return false;
            }
            if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out _)) return false;

            int offset = 13;
            OfficeColor[]? globalColorTable = null;
            byte packed = bytes[10];
            int backgroundColorIndex = bytes[11];
            if ((packed & 0x80) != 0) {
                int colorCount = 1 << ((packed & 0x07) + 1);
                if (!TryReadColorTable(bytes, ref offset, colorCount, out globalColorTable)) {
                    return false;
                }
            }
            if (validateAllFrames &&
                (globalColorTable == null
                    ? backgroundColorIndex != 0
                    : backgroundColorIndex >= globalColorTable.Length)) {
                return false;
            }

            OfficeColor backgroundColor = default;
            OfficeRasterImage? canvas = null;
            int transparentIndex = -1;
            int disposalMethod = 0;
            bool hasPendingGraphicControl = false;
            FrameRectangle previousFrame = default;
            int previousDisposalMethod = 0;
            OfficeRasterImage? restoreCanvas = null;
            long decodedFramePixels = 0;
            while (offset < bytes.Length) {
                byte marker = bytes[offset++];
                if (marker == 0x3B) {
                    return validateAllFrames
                        ? frameCount > 0 && !hasPendingGraphicControl && offset == bytes.Length
                        : image != null;
                }

                if (marker == 0x21) {
                    if (offset >= bytes.Length) {
                        return false;
                    }
                    if (validateAllFrames && signature != "GIF89a") return false;

                    byte label = bytes[offset++];
                    if (label == 0xF9) {
                        if (!TryReadGraphicControlExtension(
                            bytes,
                            ref offset,
                            out transparentIndex,
                            out disposalMethod,
                            out bool hasReservedBits)) {
                            return false;
                        }
                        if (validateAllFrames && (hasReservedBits || disposalMethod > 3)) {
                            return false;
                        }
                        if (validateAllFrames && hasPendingGraphicControl) return false;
                        hasPendingGraphicControl = true;
                    } else if (validateAllFrames && label == 0xFF) {
                        if (!TryReadFixedHeaderExtension(bytes, ref offset, expectedHeaderLength: 11)) return false;
                    } else if (label == 0x01) {
                        if (validateAllFrames) {
                            if (transparentIndex >= 0 &&
                                (globalColorTable == null || transparentIndex >= globalColorTable.Length)) return false;
                            if (!TryReadPlainTextExtension(
                                bytes,
                                ref offset,
                                width,
                                height,
                                globalColorTable)) return false;
                        } else if (!SkipSubBlocks(bytes, ref offset)) {
                            return false;
                        }
                        transparentIndex = -1;
                        disposalMethod = 0;
                        hasPendingGraphicControl = false;
                    } else if (!SkipSubBlocks(bytes, ref offset)) {
                        return false;
                    }

                    continue;
                }

                if (marker != 0x2C) {
                    return false;
                }

                if (validateAllFrames || frameCount <= frameIndex) {
                    if (canvas == null) {
                        backgroundColor = ResolveCanvasBackground(globalColorTable, backgroundColorIndex, transparentIndex);
                        canvas = new OfficeRasterImage(width, height, backgroundColor);
                    }
                    ApplyDisposal(canvas, previousFrame, previousDisposalMethod, backgroundColor, restoreCanvas);
                    restoreCanvas = disposalMethod == 3
                        ? OfficeRasterImage.FromRgba32(canvas.Width, canvas.Height, canvas.GetPixels())
                        : null;
                    if (!TryReadImageFrame(
                        bytes,
                        ref offset,
                        width,
                        height,
                        globalColorTable,
                        transparentIndex,
                        canvas,
                        requireCompleteLzw: validateAllFrames,
                        ref decodedFramePixels,
                        out FrameRectangle frame)) {
                        return false;
                    }

                    if (!validateAllFrames && frameCount == frameIndex) {
                        image = OfficeRasterImage.FromRgba32(canvas.Width, canvas.Height, canvas.GetPixels());
                    }
                    previousFrame = frame;
                    previousDisposalMethod = disposalMethod;
                } else if (!TrySkipImageFrame(
                    bytes,
                    ref offset,
                    width,
                    height,
                    globalColorTable != null)) {
                    return false;
                }
                frameCount++;
                transparentIndex = -1;
                disposalMethod = 0;
                hasPendingGraphicControl = false;
            }

            return !validateAllFrames && image != null;
        } catch {
            image = null;
            frameCount = 0;
            return false;
        }
    }

    private static bool TryReadFixedHeaderExtension(byte[] bytes, ref int offset, int expectedHeaderLength) {
        if (offset >= bytes.Length || bytes[offset++] != expectedHeaderLength ||
            offset > bytes.Length - expectedHeaderLength) {
            return false;
        }

        offset += expectedHeaderLength;
        return SkipSubBlocks(bytes, ref offset);
    }

    private static bool TryReadPlainTextExtension(
        byte[] bytes,
        ref int offset,
        int canvasWidth,
        int canvasHeight,
        OfficeColor[]? globalColorTable) {
        const int headerLength = 12;
        if (globalColorTable == null || globalColorTable.Length == 0 ||
            offset >= bytes.Length || bytes[offset++] != headerLength ||
            offset > bytes.Length - headerLength) {
            return false;
        }

        int left = ReadUInt16LittleEndian(bytes, offset);
        int top = ReadUInt16LittleEndian(bytes, offset + 2);
        int width = ReadUInt16LittleEndian(bytes, offset + 4);
        int height = ReadUInt16LittleEndian(bytes, offset + 6);
        int cellWidth = bytes[offset + 8];
        int cellHeight = bytes[offset + 9];
        int foregroundIndex = bytes[offset + 10];
        int backgroundIndex = bytes[offset + 11];
        if (width <= 0 || height <= 0 || cellWidth <= 0 || cellHeight <= 0 ||
            (long)left + width > canvasWidth || (long)top + height > canvasHeight ||
            foregroundIndex >= globalColorTable.Length || backgroundIndex >= globalColorTable.Length) {
            return false;
        }

        offset += headerLength;
        return SkipSubBlocks(bytes, ref offset);
    }

    private static bool TryReadImageFrame(
        byte[] bytes,
        ref int offset,
        int canvasWidth,
        int canvasHeight,
        OfficeColor[]? globalColorTable,
        int transparentIndex,
        OfficeRasterImage canvas,
        bool requireCompleteLzw,
        ref long decodedFramePixels,
        out FrameRectangle frame) {
        frame = default;
        if (offset + 9 > bytes.Length) {
            return false;
        }

        int left = ReadUInt16LittleEndian(bytes, offset);
        int top = ReadUInt16LittleEndian(bytes, offset + 2);
        int width = ReadUInt16LittleEndian(bytes, offset + 4);
        int height = ReadUInt16LittleEndian(bytes, offset + 6);
        byte packed = bytes[offset + 8];
        offset += 9;
        if ((requireCompleteLzw && (packed & 0x18) != 0) ||
            width <= 0 || height <= 0 || left < 0 || top < 0 ||
            left + width > canvasWidth || top + height > canvasHeight) {
            return false;
        }
        if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out int framePixels)) return false;
        if (decodedFramePixels > OfficeRasterGuards.MaximumPixels - framePixels) return false;
        decodedFramePixels += framePixels;

        OfficeColor[]? colorTable = globalColorTable;
        if ((packed & 0x80) != 0) {
            int colorCount = 1 << ((packed & 0x07) + 1);
            if (!TryReadColorTable(bytes, ref offset, colorCount, out colorTable)) {
                return false;
            }
        }

        if (colorTable == null || colorTable.Length == 0 ||
            requireCompleteLzw && transparentIndex >= colorTable.Length ||
            offset >= bytes.Length) {
            return false;
        }

        int minimumCodeSize = bytes[offset++];
        if (minimumCodeSize < 2 || minimumCodeSize > 8) {
            return false;
        }

        if (!TryReadSubBlockBytes(bytes, ref offset, out byte[] lzwBytes) ||
            !TryDecodeLzw(lzwBytes, minimumCodeSize, framePixels, requireCompleteLzw, out byte[] indices)) {
            return false;
        }

        bool interlaced = (packed & 0x40) != 0;
        int sourceIndex = 0;
        foreach (int row in EnumerateRows(height, interlaced)) {
            for (int x = 0; x < width; x++) {
                if (sourceIndex >= indices.Length) {
                    return false;
                }

                int colorIndex = indices[sourceIndex++];
                if (colorIndex >= colorTable.Length) {
                    return false;
                }

                OfficeColor color = colorTable[colorIndex];
                if (colorIndex == transparentIndex) {
                    continue;
                }

                canvas.SetPixel(left + x, top + row, color);
            }
        }

        frame = new FrameRectangle(left, top, width, height);
        return true;
    }

    private static bool TrySkipImageFrame(
        byte[] bytes,
        ref int offset,
        int canvasWidth,
        int canvasHeight,
        bool hasGlobalColorTable) {
        if (offset + 9 > bytes.Length) return false;

        int left = ReadUInt16LittleEndian(bytes, offset);
        int top = ReadUInt16LittleEndian(bytes, offset + 2);
        int width = ReadUInt16LittleEndian(bytes, offset + 4);
        int height = ReadUInt16LittleEndian(bytes, offset + 6);
        byte packed = bytes[offset + 8];
        offset += 9;
        if (width <= 0 || height <= 0 || left + width > canvasWidth || top + height > canvasHeight ||
            !OfficeRasterGuards.TryEnsurePixelCount(width, height, out _)) {
            return false;
        }

        bool hasColorTable = hasGlobalColorTable;
        if ((packed & 0x80) != 0) {
            int colorCount = 1 << ((packed & 0x07) + 1);
            int colorBytes = colorCount * 3;
            if (offset + colorBytes > bytes.Length) return false;
            offset += colorBytes;
            hasColorTable = true;
        }

        if (!hasColorTable || offset >= bytes.Length) return false;
        int minimumCodeSize = bytes[offset++];
        return minimumCodeSize >= 2 && minimumCodeSize <= 8 && SkipSubBlocks(bytes, ref offset);
    }

    private static void ApplyDisposal(OfficeRasterImage canvas, FrameRectangle frame, int disposalMethod, OfficeColor backgroundColor, OfficeRasterImage? restoreCanvas) {
        if (frame.Width <= 0 || frame.Height <= 0) return;
        if (disposalMethod == 2) {
            for (int y = frame.Top; y < frame.Top + frame.Height; y++) {
                for (int x = frame.Left; x < frame.Left + frame.Width; x++) {
                    canvas.SetPixel(x, y, backgroundColor);
                }
            }
        } else if (disposalMethod == 3 && restoreCanvas != null) {
            Buffer.BlockCopy(restoreCanvas.PixelBuffer, 0, canvas.PixelBuffer, 0, canvas.PixelBuffer.Length);
        }
    }

    private static OfficeColor ResolveCanvasBackground(OfficeColor[]? globalColorTable, int backgroundColorIndex, int transparentIndex) {
        if (globalColorTable == null ||
            backgroundColorIndex < 0 ||
            backgroundColorIndex >= globalColorTable.Length ||
            backgroundColorIndex == transparentIndex) {
            return OfficeColor.Transparent;
        }

        return globalColorTable[backgroundColorIndex];
    }

    private static bool TryDecodeLzw(
        byte[] data,
        int minimumCodeSize,
        int expectedPixelCount,
        bool requireCompleteStream,
        out byte[] indices) {
        indices = Array.Empty<byte>();
        int clearCode = 1 << minimumCodeSize;
        int endCode = clearCode + 1;
        var output = new List<byte>(expectedPixelCount);
        var dictionary = new List<byte[]>(4096);
        var reader = new LzwBitReader(data);
        int codeSize = minimumCodeSize + 1;
        int previousCode = -1;

        void ResetDictionary() {
            dictionary.Clear();
            for (int i = 0; i < clearCode; i++) {
                dictionary.Add(new[] { (byte)i });
            }

            dictionary.Add(Array.Empty<byte>());
            dictionary.Add(Array.Empty<byte>());
            codeSize = minimumCodeSize + 1;
            previousCode = -1;
        }

        ResetDictionary();
        bool sawEndCode = false;
        while (requireCompleteStream || output.Count < expectedPixelCount) {
            int code = reader.ReadBits(codeSize);
            if (code < 0) {
                return false;
            }

            if (code == clearCode) {
                ResetDictionary();
                continue;
            }

            if (code == endCode) {
                sawEndCode = true;
                break;
            }

            byte[] entry;
            if (code < dictionary.Count) {
                entry = dictionary[code];
            } else if (code == dictionary.Count && previousCode >= 0) {
                byte[] previous = dictionary[previousCode];
                entry = Append(previous, previous[0]);
            } else {
                return false;
            }

            if (requireCompleteStream && entry.Length > expectedPixelCount - output.Count) {
                return false;
            }
            output.AddRange(entry);
            if (previousCode >= 0 && dictionary.Count < 4096) {
                dictionary.Add(Append(dictionary[previousCode], entry[0]));
                if (dictionary.Count == (1 << codeSize) && codeSize < 12) {
                    codeSize++;
                }
            }

            previousCode = code;
        }

        if (requireCompleteStream && (!sawEndCode || output.Count != expectedPixelCount || !reader.HasNoTrailingBytes) ||
            !requireCompleteStream && output.Count < expectedPixelCount) {
            return false;
        }

        indices = requireCompleteStream ? output.ToArray() : output.GetRange(0, expectedPixelCount).ToArray();
        return true;
    }

    private static bool TryReadGraphicControlExtension(
        byte[] bytes,
        ref int offset,
        out int transparentIndex,
        out int disposalMethod,
        out bool hasReservedBits) {
        transparentIndex = -1;
        disposalMethod = 0;
        hasReservedBits = false;
        if (offset >= bytes.Length) {
            return false;
        }

        int blockSize = bytes[offset++];
        if (blockSize != 4 || offset + 5 > bytes.Length) {
            return false;
        }

        byte packed = bytes[offset];
        byte index = bytes[offset + 3];
        offset += 4;
        if (bytes[offset++] != 0) {
            return false;
        }

        if ((packed & 0x01) != 0) {
            transparentIndex = index;
        }

        disposalMethod = (packed >> 2) & 0x07;
        hasReservedBits = (packed & 0xE0) != 0;

        return true;
    }

    private static bool TryReadColorTable(byte[] bytes, ref int offset, int colorCount, out OfficeColor[]? colors) {
        colors = null;
        if (colorCount <= 0 || offset + (colorCount * 3) > bytes.Length) {
            return false;
        }

        colors = new OfficeColor[colorCount];
        for (int i = 0; i < colorCount; i++) {
            colors[i] = OfficeColor.FromRgb(bytes[offset], bytes[offset + 1], bytes[offset + 2]);
            offset += 3;
        }

        return true;
    }

    private static bool TryReadSubBlockBytes(byte[] bytes, ref int offset, out byte[] data) {
        data = Array.Empty<byte>();
        var buffer = new List<byte>();
        while (offset < bytes.Length) {
            int count = bytes[offset++];
            if (count == 0) {
                data = buffer.ToArray();
                return true;
            }

            if (offset + count > bytes.Length) {
                return false;
            }

            for (int i = 0; i < count; i++) {
                buffer.Add(bytes[offset + i]);
            }

            offset += count;
        }

        return false;
    }

    private static bool SkipSubBlocks(byte[] bytes, ref int offset) {
        while (offset < bytes.Length) {
            int count = bytes[offset++];
            if (count == 0) {
                return true;
            }

            if (offset + count > bytes.Length) {
                return false;
            }

            offset += count;
        }

        return false;
    }

    private static IEnumerable<int> EnumerateRows(int height, bool interlaced) {
        if (!interlaced) {
            for (int y = 0; y < height; y++) {
                yield return y;
            }

            yield break;
        }

        int[] starts = { 0, 4, 2, 1 };
        int[] steps = { 8, 8, 4, 2 };
        for (int pass = 0; pass < starts.Length; pass++) {
            for (int y = starts[pass]; y < height; y += steps[pass]) {
                yield return y;
            }
        }
    }

    private static byte[] Append(byte[] value, byte suffix) {
        byte[] result = new byte[value.Length + 1];
        Buffer.BlockCopy(value, 0, result, 0, value.Length);
        result[result.Length - 1] = suffix;
        return result;
    }

    private static int ReadUInt16LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | (bytes[offset + 1] << 8);

    private static string GetAscii(byte[] data, int offset, int count) =>
        System.Text.Encoding.ASCII.GetString(data, offset, count);

    private readonly struct FrameRectangle {
        internal FrameRectangle(int left, int top, int width, int height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        internal int Left { get; }
        internal int Top { get; }
        internal int Width { get; }
        internal int Height { get; }
    }

    private sealed class LzwBitReader {
        private readonly byte[] _data;
        private int _bitOffset;

        internal LzwBitReader(byte[] data) {
            _data = data;
        }

        internal bool HasNoTrailingBytes => (_data.Length * 8) - _bitOffset < 8;

        internal int ReadBits(int count) {
            if (count <= 0 || count > 12 || _bitOffset + count > _data.Length * 8) {
                return -1;
            }

            int value = 0;
            for (int i = 0; i < count; i++) {
                int absolute = _bitOffset + i;
                int bit = (_data[absolute / 8] >> (absolute % 8)) & 1;
                value |= bit << i;
            }

            _bitOffset += count;
            return value;
        }
    }
}

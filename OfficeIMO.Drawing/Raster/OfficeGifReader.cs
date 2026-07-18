using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free GIF decoder for the first image frame.
/// </summary>
public static class OfficeGifReader {
    /// <summary>
    /// Attempts to decode the first GIF image frame into an RGBA raster buffer.
    /// </summary>
    public static bool TryDecode(byte[]? bytes, out OfficeRasterImage? image) {
        image = null;
        try {
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

            int transparentIndex = -1;
            while (offset < bytes.Length) {
                byte marker = bytes[offset++];
                if (marker == 0x3B) {
                    return false;
                }

                if (marker == 0x21) {
                    if (offset >= bytes.Length) {
                        return false;
                    }

                    byte label = bytes[offset++];
                    if (label == 0xF9) {
                        if (!TryReadGraphicControlExtension(bytes, ref offset, out transparentIndex)) {
                            return false;
                        }
                    } else if (!SkipSubBlocks(bytes, ref offset)) {
                        return false;
                    }

                    continue;
                }

                if (marker != 0x2C) {
                    return false;
                }

                return TryReadImageFrame(bytes, ref offset, width, height, globalColorTable, backgroundColorIndex, transparentIndex, out image);
            }

            return false;
        } catch {
            image = null;
            return false;
        }
    }

    private static bool TryReadImageFrame(byte[] bytes, ref int offset, int canvasWidth, int canvasHeight, OfficeColor[]? globalColorTable, int backgroundColorIndex, int transparentIndex, out OfficeRasterImage? image) {
        image = null;
        if (offset + 9 > bytes.Length) {
            return false;
        }

        int left = ReadUInt16LittleEndian(bytes, offset);
        int top = ReadUInt16LittleEndian(bytes, offset + 2);
        int width = ReadUInt16LittleEndian(bytes, offset + 4);
        int height = ReadUInt16LittleEndian(bytes, offset + 6);
        byte packed = bytes[offset + 8];
        offset += 9;
        if (width <= 0 || height <= 0 || left < 0 || top < 0 ||
            left + width > canvasWidth || top + height > canvasHeight) {
            return false;
        }
        if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out int framePixels)) return false;

        OfficeColor[]? colorTable = globalColorTable;
        if ((packed & 0x80) != 0) {
            int colorCount = 1 << ((packed & 0x07) + 1);
            if (!TryReadColorTable(bytes, ref offset, colorCount, out colorTable)) {
                return false;
            }
        }

        if (colorTable == null || colorTable.Length == 0 || offset >= bytes.Length) {
            return false;
        }

        int minimumCodeSize = bytes[offset++];
        if (minimumCodeSize < 2 || minimumCodeSize > 8) {
            return false;
        }

        if (!TryReadSubBlockBytes(bytes, ref offset, out byte[] lzwBytes) ||
            !TryDecodeLzw(lzwBytes, minimumCodeSize, framePixels, out byte[] indices)) {
            return false;
        }

        bool interlaced = (packed & 0x40) != 0;
        OfficeColor backgroundColor = ResolveCanvasBackground(globalColorTable, backgroundColorIndex, transparentIndex);
        var result = new OfficeRasterImage(canvasWidth, canvasHeight, backgroundColor);
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
                    color = OfficeColor.FromRgba(color.R, color.G, color.B, 0);
                }

                result.SetPixel(left + x, top + row, color);
            }
        }

        image = result;
        return true;
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

    private static bool TryDecodeLzw(byte[] data, int minimumCodeSize, int expectedPixelCount, out byte[] indices) {
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
        while (output.Count < expectedPixelCount) {
            int code = reader.ReadBits(codeSize);
            if (code < 0) {
                return false;
            }

            if (code == clearCode) {
                ResetDictionary();
                continue;
            }

            if (code == endCode) {
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

            output.AddRange(entry);
            if (previousCode >= 0 && dictionary.Count < 4096) {
                dictionary.Add(Append(dictionary[previousCode], entry[0]));
                if (dictionary.Count == (1 << codeSize) && codeSize < 12) {
                    codeSize++;
                }
            }

            previousCode = code;
        }

        if (output.Count < expectedPixelCount) {
            return false;
        }

        indices = output.GetRange(0, expectedPixelCount).ToArray();
        return true;
    }

    private static bool TryReadGraphicControlExtension(byte[] bytes, ref int offset, out int transparentIndex) {
        transparentIndex = -1;
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

    private sealed class LzwBitReader {
        private readonly byte[] _data;
        private int _bitOffset;

        internal LzwBitReader(byte[] data) {
            _data = data;
        }

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

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    internal static partial class VisioPngRenderer {

        private sealed class PngRaster {
            private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
            private readonly byte[] _pixels;

            private PngRaster(int width, int height, byte[] pixels) {
                Width = width;
                Height = height;
                _pixels = pixels;
            }

            internal int Width { get; }

            internal int Height { get; }

            internal Color GetPixel(int x, int y) {
                int offset = ((y * Width) + x) * 4;
                return Color.FromRgba(_pixels[offset], _pixels[offset + 1], _pixels[offset + 2], _pixels[offset + 3]);
            }

            private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
                (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

            private static bool HasPngSignature(byte[] bytes) {
                if (bytes.Length < Signature.Length) {
                    return false;
                }

                for (int i = 0; i < Signature.Length; i++) {
                    if (bytes[i] != Signature[i]) {
                        return false;
                    }
                }

                return true;
            }

            internal static bool TryDecode(byte[] bytes, out PngRaster? image) {
                image = null;
                try {
                    if (!HasPngSignature(bytes)) {
                        return false;
                    }

                    int width = 0;
                    int height = 0;
                    int bitDepth = 0;
                    int colorType = 0;
                    int compressionMethod = 0;
                    int filterMethod = 0;
                    int interlaceMethod = 0;
                    byte[]? palette = null;
                    byte[]? transparency = null;
                    using MemoryStream idat = new();
                    int offset = Signature.Length;
                    while (offset + 12 <= bytes.Length) {
                        int length = ReadBigEndianInt32(bytes, offset);
                        if (length < 0 || offset + 12 + length > bytes.Length) {
                            return false;
                        }

                        string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
                        int dataOffset = offset + 8;
                        if (type == "IHDR") {
                            width = ReadBigEndianInt32(bytes, dataOffset);
                            height = ReadBigEndianInt32(bytes, dataOffset + 4);
                            bitDepth = bytes[dataOffset + 8];
                            colorType = bytes[dataOffset + 9];
                            compressionMethod = bytes[dataOffset + 10];
                            filterMethod = bytes[dataOffset + 11];
                            interlaceMethod = bytes[dataOffset + 12];
                        } else if (type == "PLTE") {
                            palette = new byte[length];
                            Buffer.BlockCopy(bytes, dataOffset, palette, 0, length);
                        } else if (type == "tRNS") {
                            transparency = new byte[length];
                            Buffer.BlockCopy(bytes, dataOffset, transparency, 0, length);
                        } else if (type == "IDAT") {
                            idat.Write(bytes, dataOffset, length);
                        } else if (type == "IEND") {
                            break;
                        }

                        offset = dataOffset + length + 4;
                    }

                    if (width <= 0 || height <= 0 || compressionMethod != 0 || filterMethod != 0 || interlaceMethod != 0 ||
                        !IsSupportedColorLayout(colorType, bitDepth, palette)) {
                        return false;
                    }

                    int bitsPerPixel = GetBitsPerPixel(colorType, bitDepth);
                    int bytesPerPixel = Math.Max(1, (bitsPerPixel + 7) / 8);
                    byte[] compressed = idat.ToArray();
                    if (compressed.Length < 6) {
                        return false;
                    }

                    using MemoryStream source = new(compressed, 2, compressed.Length - 6);
                    using DeflateStream deflate = new(source, CompressionMode.Decompress);
                    using MemoryStream inflated = new();
                    deflate.CopyTo(inflated);
                    byte[] scanlines = inflated.ToArray();
                    int stride = ((width * bitsPerPixel) + 7) / 8;
                    byte[] previous = new byte[stride];
                    byte[] current = new byte[stride];
                    byte[] rgba = new byte[width * height * 4];
                    int sourceOffset = 0;
                    for (int y = 0; y < height; y++) {
                        if (sourceOffset >= scanlines.Length) return false;
                        int filter = scanlines[sourceOffset++];
                        if (sourceOffset + stride > scanlines.Length) return false;
                        Buffer.BlockCopy(scanlines, sourceOffset, current, 0, stride);
                        sourceOffset += stride;
                        Unfilter(current, previous, bytesPerPixel, filter);
                        ExpandScanline(current, width, y, colorType, bitDepth, palette, transparency, rgba);

                        byte[] temp = previous;
                        previous = current;
                        current = temp;
                        Array.Clear(current, 0, current.Length);
                    }

                    image = new PngRaster(width, height, rgba);
                    return true;
                } catch {
                    image = null;
                    return false;
                }
            }

            private static bool IsSupportedColorLayout(int colorType, int bitDepth, byte[]? palette) {
                switch (colorType) {
                    case 0:
                        return bitDepth == 1 || bitDepth == 2 || bitDepth == 4 || bitDepth == 8 || bitDepth == 16;
                    case 2:
                    case 4:
                    case 6:
                        return bitDepth == 8 || bitDepth == 16;
                    case 3:
                        return (bitDepth == 1 || bitDepth == 2 || bitDepth == 4 || bitDepth == 8) &&
                               palette != null &&
                               palette.Length >= 3 &&
                               palette.Length % 3 == 0;
                    default:
                        return false;
                }
            }

            private static int GetBitsPerPixel(int colorType, int bitDepth) {
                switch (colorType) {
                    case 0:
                    case 3:
                        return bitDepth;
                    case 2:
                        return bitDepth * 3;
                    case 4:
                        return bitDepth * 2;
                    case 6:
                        return bitDepth * 4;
                    default:
                        throw new InvalidDataException("Unsupported PNG color type.");
                }
            }

            private static void ExpandScanline(
                byte[] current,
                int width,
                int y,
                int colorType,
                int bitDepth,
                byte[]? palette,
                byte[]? transparency,
                byte[] rgba) {
                for (int x = 0; x < width; x++) {
                    int targetPixel = ((y * width) + x) * 4;
                    switch (colorType) {
                        case 0:
                            ExpandGrayscale(
                                GetGrayscaleSample(current, x, bitDepth),
                                bitDepth,
                                transparency,
                                rgba,
                                targetPixel);
                            break;
                        case 2:
                            ExpandTrueColor(current, x * (bitDepth == 16 ? 6 : 3), bitDepth, transparency, rgba, targetPixel);
                            break;
                        case 3:
                            ExpandPalette(GetPackedSample(current, x, bitDepth), palette!, transparency, rgba, targetPixel);
                            break;
                        case 4:
                            ExpandGrayscaleAlpha(current, x * (bitDepth == 16 ? 4 : 2), bitDepth, rgba, targetPixel);
                            break;
                        case 6:
                            ExpandTrueColorAlpha(current, x * (bitDepth == 16 ? 8 : 4), bitDepth, rgba, targetPixel);
                            break;
                        default:
                            throw new InvalidDataException("Unsupported PNG color type.");
                    }
                }
            }

            private static void ExpandGrayscale(int sample, int bitDepth, byte[]? transparency, byte[] rgba, int targetPixel) {
                byte gray = ScaleSample(sample, bitDepth);
                rgba[targetPixel] = gray;
                rgba[targetPixel + 1] = gray;
                rgba[targetPixel + 2] = gray;
                rgba[targetPixel + 3] = IsTransparentGray(sample, transparency) ? (byte)0 : (byte)255;
            }

            private static void ExpandGrayscaleAlpha(byte[] current, int sourcePixel, int bitDepth, byte[] rgba, int targetPixel) {
                int graySample = bitDepth == 16 ? ReadBigEndianUInt16(current, sourcePixel) : current[sourcePixel];
                int alphaSample = bitDepth == 16 ? ReadBigEndianUInt16(current, sourcePixel + 2) : current[sourcePixel + 1];
                byte gray = ScaleSample(graySample, bitDepth);
                rgba[targetPixel] = gray;
                rgba[targetPixel + 1] = gray;
                rgba[targetPixel + 2] = gray;
                rgba[targetPixel + 3] = ScaleSample(alphaSample, bitDepth);
            }

            private static void ExpandTrueColor(byte[] current, int sourcePixel, int bitDepth, byte[]? transparency, byte[] rgba, int targetPixel) {
                int redSample;
                int greenSample;
                int blueSample;
                if (bitDepth == 16) {
                    redSample = ReadBigEndianUInt16(current, sourcePixel);
                    greenSample = ReadBigEndianUInt16(current, sourcePixel + 2);
                    blueSample = ReadBigEndianUInt16(current, sourcePixel + 4);
                } else {
                    redSample = current[sourcePixel];
                    greenSample = current[sourcePixel + 1];
                    blueSample = current[sourcePixel + 2];
                }

                rgba[targetPixel] = ScaleSample(redSample, bitDepth);
                rgba[targetPixel + 1] = ScaleSample(greenSample, bitDepth);
                rgba[targetPixel + 2] = ScaleSample(blueSample, bitDepth);
                rgba[targetPixel + 3] = IsTransparentRgb(redSample, greenSample, blueSample, transparency) ? (byte)0 : (byte)255;
            }

            private static void ExpandTrueColorAlpha(byte[] current, int sourcePixel, int bitDepth, byte[] rgba, int targetPixel) {
                if (bitDepth == 16) {
                    rgba[targetPixel] = ScaleSample(ReadBigEndianUInt16(current, sourcePixel), bitDepth);
                    rgba[targetPixel + 1] = ScaleSample(ReadBigEndianUInt16(current, sourcePixel + 2), bitDepth);
                    rgba[targetPixel + 2] = ScaleSample(ReadBigEndianUInt16(current, sourcePixel + 4), bitDepth);
                    rgba[targetPixel + 3] = ScaleSample(ReadBigEndianUInt16(current, sourcePixel + 6), bitDepth);
                } else {
                    Buffer.BlockCopy(current, sourcePixel, rgba, targetPixel, 4);
                }
            }

            private static void ExpandPalette(int index, byte[] palette, byte[]? transparency, byte[] rgba, int targetPixel) {
                int paletteOffset = index * 3;
                if (paletteOffset + 2 >= palette.Length) {
                    throw new InvalidDataException("PNG palette index is outside PLTE.");
                }

                rgba[targetPixel] = palette[paletteOffset];
                rgba[targetPixel + 1] = palette[paletteOffset + 1];
                rgba[targetPixel + 2] = palette[paletteOffset + 2];
                rgba[targetPixel + 3] = transparency != null && index < transparency.Length ? transparency[index] : (byte)255;
            }

            private static int GetPackedSample(byte[] current, int x, int bitDepth) {
                if (bitDepth == 8) {
                    return current[x];
                }

                int samplesPerByte = 8 / bitDepth;
                int shift = (samplesPerByte - 1 - (x % samplesPerByte)) * bitDepth;
                int mask = (1 << bitDepth) - 1;
                return (current[x / samplesPerByte] >> shift) & mask;
            }

            private static int GetGrayscaleSample(byte[] current, int x, int bitDepth) {
                if (bitDepth == 16) {
                    return ReadBigEndianUInt16(current, x * 2);
                }

                return bitDepth == 8 ? current[x] : GetPackedSample(current, x, bitDepth);
            }

            private static int ReadBigEndianUInt16(byte[] bytes, int offset) =>
                (bytes[offset] << 8) | bytes[offset + 1];

            private static byte ScaleSample(int sample, int bitDepth) {
                if (bitDepth == 8) {
                    return (byte)sample;
                }

                int max = (1 << bitDepth) - 1;
                return (byte)Math.Round(sample * 255D / max);
            }

            private static bool IsTransparentGray(int sample, byte[]? transparency) =>
                transparency != null && transparency.Length >= 2 && sample == ((transparency[0] << 8) | transparency[1]);

            private static bool IsTransparentRgb(int red, int green, int blue, byte[]? transparency) =>
                transparency != null &&
                transparency.Length >= 6 &&
                red == ((transparency[0] << 8) | transparency[1]) &&
                green == ((transparency[2] << 8) | transparency[3]) &&
                blue == ((transparency[4] << 8) | transparency[5]);

            private static void Unfilter(byte[] current, byte[] previous, int bytesPerPixel, int filter) {
                for (int i = 0; i < current.Length; i++) {
                    int left = i >= bytesPerPixel ? current[i - bytesPerPixel] : 0;
                    int up = previous[i];
                    int upLeft = i >= bytesPerPixel ? previous[i - bytesPerPixel] : 0;
                    int value = current[i];
                    switch (filter) {
                        case 0:
                            break;
                        case 1:
                            value += left;
                            break;
                        case 2:
                            value += up;
                            break;
                        case 3:
                            value += (left + up) / 2;
                            break;
                        case 4:
                            value += Paeth(left, up, upLeft);
                            break;
                        default:
                            throw new InvalidDataException("Unsupported PNG filter.");
                    }

                    current[i] = (byte)(value & 0xFF);
                }
            }

            private static int Paeth(int left, int up, int upLeft) {
                int p = left + up - upLeft;
                int pa = Math.Abs(p - left);
                int pb = Math.Abs(p - up);
                int pc = Math.Abs(p - upLeft);
                if (pa <= pb && pa <= pc) return left;
                return pb <= pc ? up : upLeft;
            }
        }
    }
}

using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    /// <summary>
    /// Attempts to decode classic grayscale, RGB, RGBA, or device-CMYK TIFF with unsigned 8/16-bit or finite floating 16/24/32/64-bit samples, or eight-bit palette TIFF using
    /// chunky or planar strips or tiles with uncompressed, LZW, PackBits, or Deflate payloads.
    /// Floating samples are normalized device components; JPEG-compressed and BigTIFF payloads remain caller-codec responsibilities.
    /// </summary>
    public static bool TryDecode(byte[]? encodedBytes, out OfficeRasterImage? image) =>
        TryDecodePage(encodedBytes, 0, options: null, out image);

    /// <summary>Attempts to decode one zero-based page from a bounded classic TIFF container.</summary>
    public static bool TryDecodePage(byte[]? encodedBytes, int pageIndex, out OfficeRasterImage? image) =>
        TryDecodePage(encodedBytes, pageIndex, options: null, out image);

    internal static bool TryDecodePage(
        byte[]? encodedBytes,
        int pageIndex,
        OfficeRasterDecodeOptions? options,
        out OfficeRasterImage? image,
        OfficeIccColorProfile? colorProfile = null,
        OfficeIccRenderingIntent renderingIntent = OfficeIccRenderingIntent.RelativeColorimetric) {
        image = null;
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        OfficeRasterDecodeOptions effective = options ?? new OfficeRasterDecodeOptions();
        effective.Validate();
        effective.CancellationToken.ThrowIfCancellationRequested();
        if (!IsTiff(encodedBytes) || encodedBytes == null ||
            encodedBytes.Length > effective.MaximumEncodedBytes ||
            !OfficeTiffStructureValidator.TryValidate(
                encodedBytes, 0, encodedBytes.Length, effective.CancellationToken)) {
            return false;
        }
        try {
            bool littleEndian = encodedBytes[0] == (byte)'I';
            if (ReadUInt16(encodedBytes, 2, littleEndian) != 42) return false;
            int ifdOffset = ReadOffset(encodedBytes, 4, littleEndian);
            var visitedIfds = new System.Collections.Generic.HashSet<int>();
            int currentPageIndex = 0;
            while (ifdOffset != 0) {
                effective.CancellationToken.ThrowIfCancellationRequested();
                if (visitedIfds.Count >= MaximumIfdCount || !visitedIfds.Add(ifdOffset) ||
                    !HasBytes(encodedBytes, ifdOffset, 2)) {
                    return false;
                }
                int entryCount = ReadUInt16(encodedBytes, ifdOffset, littleEndian);
                if (entryCount <= 0 || !HasBytes(encodedBytes, ifdOffset + 2, checked(entryCount * 12 + 4))) return false;

                var entries = new System.Collections.Generic.Dictionary<int, TiffEntry>();
                int entryOffset = ifdOffset + 2;
                for (int index = 0; index < entryCount; index++, entryOffset += 12) {
                    if ((index & 0xFF) == 0) effective.CancellationToken.ThrowIfCancellationRequested();
                    int tag = ReadUInt16(encodedBytes, entryOffset, littleEndian);
                    int type = ReadUInt16(encodedBytes, entryOffset + 2, littleEndian);
                    uint count = ReadUInt32(encodedBytes, entryOffset + 4, littleEndian);
                    if (count == 0 || count > int.MaxValue || entries.ContainsKey(tag) ||
                        !HasValidEntryValueRange(
                            encodedBytes,
                            type,
                            (int)count,
                            entryOffset + 8,
                            littleEndian)) {
                        return false;
                    }
                    entries.Add(tag, new TiffEntry(type, (int)count, entryOffset + 8));
                }

                int nextIfdPointerOffset = checked(ifdOffset + 2 + entryCount * 12);
                int nextIfdOffset = ReadOffset(encodedBytes, nextIfdPointerOffset, littleEndian);
                if (currentPageIndex != pageIndex) {
                    ifdOffset = nextIfdOffset;
                    currentPageIndex++;
                    continue;
                }

                if (!TryReadScalar(encodedBytes, entries, 256, littleEndian, out int width) ||
                    !TryReadScalar(encodedBytes, entries, 257, littleEndian, out int height) ||
                    !IsWithinPixelLimit(width, height, effective.MaximumDecodedPixels)) {
                    return false;
                }

                if (!TryReadScalarOrDefault(encodedBytes, entries, 259, littleEndian, 1, out int compression) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 262, littleEndian, 2, out int photometric) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 278, littleEndian, height, out int rowsPerStrip) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 284, littleEndian, 1, out int planarConfiguration) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 317, littleEndian, 1, out int predictor)) {
                    return false;
                }
                int orientation = 1;
                if (!effective.IgnoreTiffOrientation &&
                    !TryReadScalarOrDefault(encodedBytes, entries, 274, littleEndian, 1, out orientation)) return false;
                if (!TryGetBaseSampleCount(photometric, out int baseSamples) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 277, littleEndian, baseSamples, out int samples)) {
                    return false;
                }
                if (photometric == 5 &&
                    (!TryReadScalarOrDefault(encodedBytes, entries, 332, littleEndian, 1, out int inkSet) ||
                     inkSet != 1)) {
                    return false;
                }
                if ((compression != (int)OfficeTiffCompression.None &&
                     compression != (int)OfficeTiffCompression.Lzw &&
                     compression != (int)OfficeTiffCompression.PackBits &&
                     compression != (int)OfficeTiffCompression.Deflate &&
                     compression != 32946) ||
                    orientation < 1 || orientation > 8 ||
                    (samples != baseSamples && samples != baseSamples + 1) ||
                    rowsPerStrip < 1 ||
                    (planarConfiguration != 1 && planarConfiguration != 2) ||
                    (predictor < 1 || predictor > 3)) {
                    return false;
                }

                if (!TryGetSampleByteCount(encodedBytes, entries, littleEndian, samples, photometric, out int sampleBytes, out bool floating) ||
                    !IsSupportedSamplePredictor(predictor, floating, compression)) {
                    return false;
                }

                int[]? colorMap = null;
                if (photometric == 3 &&
                    !TryReadValues(encodedBytes, entries, 320, littleEndian, 768, out colorMap)) {
                    return false;
                }

                if (colorProfile != null &&
                    !((photometric == 0 || photometric == 1) && colorProfile.ComponentCount == 1 ||
                      (photometric == 2 || photometric == 3) && colorProfile.ComponentCount == 3 ||
                      photometric == 5 && colorProfile.ComponentCount == 4)) return false;
                double[]? colorComponents = colorProfile == null ? null : new double[colorProfile.ComponentCount];
                int alphaKind = 2;
                if (samples == baseSamples + 1) {
                    if (!TryReadValues(encodedBytes, entries, 338, littleEndian, 1, out int[] extraSamples) ||
                        (extraSamples[0] < 0 || extraSamples[0] > 2)) {
                        return false;
                    }
                    alphaKind = extraSamples[0];
                }

                long maximumDecodeWorkBytes = OfficeRasterGuards.MaximumDecodedBytes - effective.RetainedManagedBytes;
                if (maximumDecodeWorkBytes < 1L) return false;
                var decodeWorkBudget = new TiffValidationBudget(maximumDecodeWorkBytes);
                if (!TryDecodePixelSegments(encodedBytes, entries, littleEndian, width, height, samples, sampleBytes,
                        compression, planarConfiguration, predictor, floating, alphaKind == 0 ? baseSamples : samples, effective, decodeWorkBudget,
                        retainPixels: true, out byte[] source)) return false;

                int orientedWidth = orientation >= 5 ? height : width;
                int orientedHeight = orientation >= 5 ? width : height;
                byte[] rgba = OfficeRasterGuards.AllocateRgba32(orientedWidth, orientedHeight, "TIFF decoded pixels exceed the managed limit.");
                for (int y = 0; y < height; y++) {
                    if ((y & 31) == 0) effective.CancellationToken.ThrowIfCancellationRequested();
                    for (int x = 0; x < width; x++) {
                        if ((x & 0xFFF) == 0) effective.CancellationToken.ThrowIfCancellationRequested();
                        int sourcePixel = ((y * width) + x) * samples * sampleBytes;
                        ResolveOrientedPixel(x, y, width, height, orientation, out int targetX, out int targetY);
                        int targetPixel = ((targetY * orientedWidth) + targetX) * 4;
                        byte red, green, blue, alpha;
                        if (floating) {
                            ConvertFloatingPixel(source, sourcePixel, sampleBytes, littleEndian, photometric,
                                samples > baseSamples, alphaKind, colorComponents,
                                out red, out green, out blue, out alpha);
                        } else if (sampleBytes == 2) {
                            ConvertUnsigned16Pixel(source, sourcePixel, littleEndian, photometric,
                                samples > baseSamples, alphaKind, colorComponents,
                                out red, out green, out blue, out alpha);
                        } else {
                            alpha = samples == baseSamples + 1 && alphaKind != 0
                                ? source[sourcePixel + baseSamples]
                                : (byte)255;
                            ConvertPixel(source, sourcePixel, photometric, alphaKind, alpha, colorMap,
                                out red, out green, out blue);
                            if (colorProfile != null) {
                                if (photometric == 5) {
                                    for (int channel = 0; channel < 4; channel++) {
                                        byte sample = source[sourcePixel + channel];
                                        colorComponents![channel] = (alphaKind == 1 ? Unpremultiply(sample, alpha) : sample) / 255D;
                                    }
                                } else {
                                    colorComponents![0] = red / 255D;
                                    if (colorComponents.Length == 3) { colorComponents[1] = green / 255D; colorComponents[2] = blue / 255D; }
                                }
                            }
                        }
                        if (colorProfile != null) {
                            if (!colorProfile.TryConvert(colorComponents!, renderingIntent, out OfficeColor converted)) return false;
                            red = converted.R; green = converted.G; blue = converted.B;
                        }
                        rgba[targetPixel] = red;
                        rgba[targetPixel + 1] = green;
                        rgba[targetPixel + 2] = blue;
                        rgba[targetPixel + 3] = alpha;
                    }
                }
                image = OfficeRasterImage.FromOwnedRgba32(orientedWidth, orientedHeight, rgba);
                return true;
            }
            return false;
        } catch (ArgumentException) {
            return false;
        } catch (FormatException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
    }

}

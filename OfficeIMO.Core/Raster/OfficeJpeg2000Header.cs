using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Recognizes the bounded, opaque Gray/RGB subset safe for codec pass-through.</summary>
internal static class OfficeJpeg2000Header {
    internal static bool TryGetOpaqueComponents(byte[] bytes, out int components) {
        return TryGetOpaqueDimensions(bytes, out components, out _, out _);
    }

    internal static bool TryGetOpaqueDimensions(byte[] bytes, out int components, out int width, out int height) {
        return TryGetOpaqueDimensions(bytes, CancellationToken.None, out components, out width, out height);
    }

    internal static bool TryGetOpaqueDimensions(
        byte[] bytes,
        CancellationToken cancellationToken,
        out int components,
        out int width,
        out int height) {
        return TryReadOpaquePayload(
            bytes,
            requireCompleteCodestream: false,
            cancellationToken,
            out components,
            out width,
            out height);
    }

    internal static bool TryValidateOpaquePayload(byte[] bytes, out int components, out int width, out int height) {
        return TryValidateOpaquePayload(bytes, CancellationToken.None, out components, out width, out height);
    }

    internal static bool TryValidateOpaquePayload(
        byte[] bytes,
        CancellationToken cancellationToken,
        out int components,
        out int width,
        out int height) {
        return TryReadOpaquePayload(
            bytes,
            requireCompleteCodestream: true,
            cancellationToken,
            out components,
            out width,
            out height);
    }

    internal static bool IsJp2Container(byte[] bytes) =>
        bytes.Length >= 12 && Read32(bytes, 0) == 12 && Read32(bytes, 4) == 0x6A502020 &&
        Read32(bytes, 8) == 0x0D0A870A;

    private static bool TryReadOpaquePayload(
        byte[] bytes,
        bool requireCompleteCodestream,
        CancellationToken cancellationToken,
        out int components,
        out int width,
        out int height) {
        components = width = height = 0;
        cancellationToken.ThrowIfCancellationRequested();
        if (TryReadCodestream(
                bytes,
                0,
                bytes.Length,
                requireCompleteCodestream,
                cancellationToken,
                out components,
                out width,
                out height,
                out _)) return true;
        if (!IsJp2Container(bytes)) return false;
        int offset = 12;
        cancellationToken.ThrowIfCancellationRequested();
        if (!TryReadBox(bytes, ref offset, bytes.Length, out uint fileType, out int fileTypeStart, out int fileTypeEnd) ||
            fileType != 0x66747970 ||
            !TryValidateFileTypeBox(bytes, fileTypeStart, fileTypeEnd, cancellationToken)) return false;
        bool header = false, codestream = false;
        int expected = 0, expectedWidth = 0, expectedHeight = 0;
        byte[] expectedComponentPrecisions = System.Array.Empty<byte>();
        byte[] componentPrecisions = System.Array.Empty<byte>();
        while (offset < bytes.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryReadBox(bytes, ref offset, bytes.Length, out uint type, out int start, out int end)) return false;
            if (type == 0x6A703268) { // jp2h
                if (header || codestream || !TryReadHeader(
                        bytes,
                        start,
                        end,
                        cancellationToken,
                        out expected,
                        out expectedWidth,
                        out expectedHeight,
                        out expectedComponentPrecisions)) return false;
                header = true;
            } else if (type == 0x6A703263) { // jp2c
                if (!header || codestream || !TryReadCodestream(
                        bytes,
                        start,
                        end,
                        requireCompleteCodestream,
                        cancellationToken,
                        out components,
                        out width,
                        out height,
                        out componentPrecisions)) return false;
                codestream = true;
            } else if (type != 0x66726565 && type != 0x786D6C20 && type != 0x75756964) {
                // Extended JPX composition/channel metadata is outside this opaque subset.
                return false;
            }
        }
        return header && codestream && expected == components && expectedWidth == width && expectedHeight == height &&
            HaveSameComponentPrecisions(expectedComponentPrecisions, componentPrecisions);
    }

    private static bool TryValidateFileTypeBox(
        byte[] bytes,
        int start,
        int end,
        CancellationToken cancellationToken) {
        int contentLength = end - start;
        if (contentLength < 12 || contentLength % 4 != 0 || Read32(bytes, start) != 0x6A703220) return false;
        bool declaresJp2Compatibility = false;
        for (int offset = start + 8; offset < end; offset += 4) {
            cancellationToken.ThrowIfCancellationRequested();
            if (Read32(bytes, offset) == 0x6A703220) declaresJp2Compatibility = true;
        }
        return declaresJp2Compatibility;
    }

    private static bool TryReadHeader(
        byte[] bytes,
        int offset,
        int end,
        CancellationToken cancellationToken,
        out int components,
        out int width,
        out int height,
        out byte[] componentPrecisions) {
        components = width = height = 0;
        componentPrecisions = System.Array.Empty<byte>();
        int colorComponents = 0;
        int bitsPerComponent = -1;
        byte[]? variableComponentPrecisions = null;
        bool hasBitsPerComponentBox = false;
        bool hasBox = false;
        while (offset < end) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryReadBox(bytes, ref offset, end, out uint type, out int start, out int boxEnd)) return false;
            if (!hasBox && type != 0x69686472) return false;
            hasBox = true;
            if (type == 0x69686472) { // ihdr
                if (components != 0 || boxEnd - start != 14) return false;
                components = Read16(bytes, start + 8);
                if (components is not (1 or 3)) return false;
                bitsPerComponent = bytes[start + 10];
                if ((bitsPerComponent != 255 && (bitsPerComponent & 0x7F) > 37) ||
                    bytes[start + 11] != 7 || // Compression type is always JPEG 2000.
                    bytes[start + 12] > 1 || // Unknown-colourspace flag.
                    bytes[start + 13] != 0) return false; // Intellectual-property metadata is outside this opaque subset.
                if (!TryBoundDimensions(Read32(bytes, start + 4), Read32(bytes, start), out width, out height)) return false;
            } else if (type == 0x636F6C72) { // colr: baseline enumerated Gray or sRGB only
                if (colorComponents != 0 || boxEnd - start != 7 ||
                    bytes[start] != 1 || // Enumerated colourspace method.
                    bytes[start + 1] != 0 || // Baseline JP2 precedence.
                    bytes[start + 2] != 0) return false; // Exact baseline approximation.
                uint color = Read32(bytes, start + 3);
                colorComponents = color == 16 ? 3 : color == 17 ? 1 : 0;
                if (colorComponents == 0) return false;
            } else if (type == 0x63646566) { // cdef: any opacity or custom association requires normalization
                if (boxEnd - start < 2) return false;
                int count = Read16(bytes, start);
                if (boxEnd - start != 2 + count * 6) return false;
                for (int i = 0; i < count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int entry = start + 2 + i * 6;
                    if (Read16(bytes, entry) != i || Read16(bytes, entry + 2) != 0 ||
                        Read16(bytes, entry + 4) != i + 1) return false;
                }
                if (count != components) return false;
            } else if (type == 0x62706363) { // bpcc
                if (bitsPerComponent != 255 || hasBitsPerComponentBox || components == 0 || boxEnd - start != components) return false;
                for (int i = 0; i < components; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if ((bytes[start + i] & 0x7F) > 37) return false;
                }
                variableComponentPrecisions = new byte[components];
                System.Buffer.BlockCopy(bytes, start, variableComponentPrecisions, 0, components);
                hasBitsPerComponentBox = true;
            } else if (type != 0x72657320) { // res
                return false; // Palette/channel remapping and extended headers are not pass-through safe.
            }
        }
        if (components != colorComponents || components is not (1 or 3) ||
            (bitsPerComponent == 255) != hasBitsPerComponentBox) return false;
        if (bitsPerComponent == 255) {
            componentPrecisions = variableComponentPrecisions!;
        } else {
            componentPrecisions = new byte[components];
            for (int i = 0; i < components; i++) componentPrecisions[i] = (byte)bitsPerComponent;
        }
        return true;
    }

    private static bool TryReadCodestream(
        byte[] bytes,
        int start,
        int end,
        bool requireCompleteCodestream,
        CancellationToken cancellationToken,
        out int components,
        out int width,
        out int height,
        out byte[] componentPrecisions) {
        components = width = height = 0;
        componentPrecisions = System.Array.Empty<byte>();
        cancellationToken.ThrowIfCancellationRequested();
        if (end - start < 42 || Read32(bytes, start) != 0xFF4FFF51) return false; // SOC, SIZ
        components = Read16(bytes, start + 40);
        int length = Read16(bytes, start + 4);
        if (components is not (1 or 3) || length != 38 + components * 3 || length > end - start - 4) return false;
        uint right = Read32(bytes, start + 8), bottom = Read32(bytes, start + 12);
        uint left = Read32(bytes, start + 16), top = Read32(bytes, start + 20);
        uint tileWidth = Read32(bytes, start + 24), tileHeight = Read32(bytes, start + 28);
        uint tileLeft = Read32(bytes, start + 32), tileTop = Read32(bytes, start + 36);
        if (right <= left || bottom <= top || tileLeft > left || tileTop > top ||
            (ulong)tileLeft + tileWidth <= left || (ulong)tileTop + tileHeight <= top ||
            !TryBoundDimensions(right, bottom, out _, out _) ||
            !TryBoundDimensions(tileWidth, tileHeight, out _, out _) ||
            !TryBoundDimensions(right - left, bottom - top, out width, out height)) return false;
        ulong tilesAcross = ((ulong)right - tileLeft + tileWidth - 1UL) / tileWidth;
        ulong tilesDown = ((ulong)bottom - tileTop + tileHeight - 1UL) / tileHeight;
        ulong tileCountValue = tilesAcross * tilesDown;
        if (tileCountValue == 0 || tileCountValue > ushort.MaxValue) return false;
        var parsedComponentPrecisions = new byte[components];
        for (int i = 0; i < components; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            int component = start + 42 + i * 3;
            if ((bytes[component] & 127) > 15 || bytes[component + 1] == 0 || bytes[component + 2] == 0) return false;
            parsedComponentPrecisions[i] = bytes[component];
        }
        if (requireCompleteCodestream && !HasCompleteCodestream(
            bytes,
            start + 4 + length,
            end,
            (int)tileCountValue,
            cancellationToken)) return false;
        componentPrecisions = parsedComponentPrecisions;
        return true;
    }

    private static bool HasCompleteCodestream(
        byte[] bytes,
        int markerOffset,
        int end,
        int tileCount,
        CancellationToken cancellationToken) {
        if (end - markerOffset < 16 || bytes[end - 2] != 0xFF || bytes[end - 1] != 0xD9) return false;

        int offset = markerOffset;
        bool hasCodingStyleDefault = false;
        bool hasQuantizationDefault = false;
        while (offset < end - 2 && !IsMarker(bytes, offset, 0x90)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryGetMarkerCode(bytes, offset, end - 2, out byte marker, out int segmentOffset)) return false;
            if (!TrySkipMarkerSegment(bytes, ref offset, end - 2, cancellationToken)) return false;
            if (marker == 0x52) {
                if (hasCodingStyleDefault || !TryValidateCodingStyleDefault(
                        bytes,
                        segmentOffset,
                        end - 2)) return false;
                hasCodingStyleDefault = true;
            } else if (marker == 0x5C) {
                if (hasQuantizationDefault || !TryValidateQuantizationDefault(
                        bytes,
                        segmentOffset,
                        end - 2)) return false;
                hasQuantizationDefault = true;
            }
        }
        if (!hasCodingStyleDefault || !hasQuantizationDefault) return false;

        bool foundTilePart = false;
        var nextPartNumbers = new int[tileCount];
        var declaredPartCounts = new int[tileCount];
        var seenTiles = new bool[tileCount];
        while (offset < end - 2) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!IsMarker(bytes, offset, 0x90) || end - offset < 14 || Read16(bytes, offset + 2) != 10) {
                return false;
            }

            int tileIndex = Read16(bytes, offset + 4);
            int partIndex = bytes[offset + 10];
            int partCount = bytes[offset + 11];
            if (tileIndex >= tileCount || partIndex != nextPartNumbers[tileIndex]) return false;
            if (declaredPartCounts[tileIndex] != 0) {
                if (partCount != declaredPartCounts[tileIndex]) return false;
            } else if (partCount != 0) {
                if (partIndex >= partCount) return false;
                declaredPartCounts[tileIndex] = partCount;
            }
            nextPartNumbers[tileIndex]++;
            seenTiles[tileIndex] = true;

            uint declaredLength = Read32(bytes, offset + 6);
            int tilePartEnd;
            if (declaredLength == 0) {
                tilePartEnd = end - 2;
            } else {
                if (declaredLength > int.MaxValue || declaredLength > end - offset - 2) return false;
                tilePartEnd = offset + (int)declaredLength;
            }

            int tileHeaderOffset = offset + 12;
            bool hasTileCodingStyleDefault = false;
            bool hasTileQuantizationDefault = false;
            while (tileHeaderOffset < tilePartEnd && !IsMarker(bytes, tileHeaderOffset, 0x93)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!TryGetMarkerCode(
                        bytes,
                        tileHeaderOffset,
                        tilePartEnd,
                        out byte tileMarker,
                        out int tileSegmentOffset)) return false;
                if (tileMarker == 0x52) {
                    if (hasTileCodingStyleDefault || !TryValidateCodingStyleDefault(
                            bytes,
                            tileSegmentOffset,
                            tilePartEnd)) return false;
                    hasTileCodingStyleDefault = true;
                } else if (tileMarker == 0x5C) {
                    if (hasTileQuantizationDefault || !TryValidateQuantizationDefault(
                            bytes,
                            tileSegmentOffset,
                            tilePartEnd)) return false;
                    hasTileQuantizationDefault = true;
                }
                if (!TrySkipMarkerSegment(bytes, ref tileHeaderOffset, tilePartEnd, cancellationToken)) return false;
            }
            if (!IsMarker(bytes, tileHeaderOffset, 0x93) || tileHeaderOffset + 2 >= tilePartEnd) return false;

            foundTilePart = true;
            offset = tilePartEnd;
            if (declaredLength == 0) break;
        }
        if (!foundTilePart || offset != end - 2) return false;
        for (int tileIndex = 0; tileIndex < tileCount; tileIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!seenTiles[tileIndex]) return false;
            if (declaredPartCounts[tileIndex] != 0 &&
                nextPartNumbers[tileIndex] != declaredPartCounts[tileIndex]) return false;
        }
        return true;
    }

    private static bool TryValidateCodingStyleDefault(
        byte[] bytes,
        int markerOffset,
        int limit) {
        if (limit - markerOffset < 14 || !IsMarker(bytes, markerOffset, 0x52)) return false;
        int length = Read16(bytes, markerOffset + 2);
        if (length < 12 || length > limit - markerOffset - 2) return false;

        int content = markerOffset + 4;
        byte codingStyle = bytes[content];
        int levels = bytes[content + 5];
        if ((codingStyle & 0xF8) != 0 ||
            bytes[content + 1] > 4 ||
            Read16(bytes, content + 2) == 0 ||
            bytes[content + 4] > 1 ||
            levels > 32 ||
            bytes[content + 6] > 8 ||
            bytes[content + 7] > 8 ||
            bytes[content + 6] + bytes[content + 7] > 8 ||
            (bytes[content + 8] & 0xC0) != 0 ||
            bytes[content + 9] > 1) return false;
        int precinctBytes = (codingStyle & 0x01) != 0 ? levels + 1 : 0;
        if (length != 12 + precinctBytes) return false;
        return true;
    }

    private static bool TryValidateQuantizationDefault(
        byte[] bytes,
        int markerOffset,
        int limit) {
        if (limit - markerOffset < 6 || !IsMarker(bytes, markerOffset, 0x5C)) return false;
        int length = Read16(bytes, markerOffset + 2);
        if (length < 4 || length > 197 || length > limit - markerOffset - 2) return false;

        int content = markerOffset + 4;
        int style = bytes[content] & 0x1F;
        int stepBytes = length - 3;
        if (style == 0) {
            return (stepBytes - 1) % 3 == 0;
        }
        if (style == 1) return stepBytes == 2;
        if (style != 2 || stepBytes < 2 || (stepBytes & 1) != 0) return false;
        int subbands = stepBytes / 2;
        return (subbands - 1) % 3 == 0;
    }

    private static bool TrySkipMarkerSegment(
        byte[] bytes,
        ref int offset,
        int limit,
        CancellationToken cancellationToken) {
        if (limit - offset < 4 || bytes[offset] != 0xFF) return false;
        int markerOffset = offset;
        while (markerOffset + 1 < limit && bytes[markerOffset + 1] == 0xFF) {
            cancellationToken.ThrowIfCancellationRequested();
            markerOffset++;
        }
        if (markerOffset + 3 >= limit || bytes[markerOffset + 1] <= 0x01 ||
            bytes[markerOffset + 1] is 0x90 or 0x93 or 0xD9) return false;
        int length = Read16(bytes, markerOffset + 2);
        if (length < 2 || length > limit - markerOffset - 2) return false;
        offset = markerOffset + 2 + length;
        return true;
    }

    private static bool IsMarker(byte[] bytes, int offset, byte marker) =>
        offset >= 0 && offset + 1 < bytes.Length && bytes[offset] == 0xFF && bytes[offset + 1] == marker;

    private static bool TryGetMarkerCode(byte[] bytes, int offset, int limit, out byte marker, out int markerOffset) {
        marker = 0;
        markerOffset = offset;
        if (offset >= limit || bytes[offset] != 0xFF) return false;
        while (offset + 1 < limit && bytes[offset + 1] == 0xFF) offset++;
        if (offset + 1 >= limit) return false;
        marker = bytes[offset + 1];
        markerOffset = offset;
        return true;
    }

    private static bool HaveSameComponentPrecisions(byte[] expected, byte[] actual) {
        if (expected.Length != actual.Length) return false;
        for (int i = 0; i < expected.Length; i++) {
            if (expected[i] != actual[i]) return false;
        }
        return true;
    }

    private static bool TryBoundDimensions(uint rawWidth, uint rawHeight, out int width, out int height) {
        width = height = 0;
        if (rawWidth == 0 || rawHeight == 0 || rawWidth > int.MaxValue || rawHeight > int.MaxValue ||
            (ulong)rawWidth * rawHeight > (ulong)OfficeRasterGuards.MaximumPixels ||
            (ulong)rawWidth * rawHeight * 4UL > (ulong)OfficeRasterGuards.MaximumDecodedBytes) return false;
        width = (int)rawWidth; height = (int)rawHeight;
        return true;
    }

    private static bool TryReadBox(byte[] bytes, ref int offset, int limit, out uint type, out int start, out int end) {
        type = 0; start = end = 0;
        if (limit - offset < 8) return false;
        uint length = Read32(bytes, offset);
        type = Read32(bytes, offset + 4);
        int header = 8;
        if (length == 1) {
            if (limit - offset < 16 || Read32(bytes, offset + 8) != 0) return false;
            length = Read32(bytes, offset + 12);
            header = 16;
        } else if (length == 0) length = (uint)(limit - offset);
        if (length < header || length > limit - offset) return false;
        start = offset + header;
        end = offset + (int)length;
        offset = end;
        return true;
    }

    private static int Read16(byte[] bytes, int offset) => (bytes[offset] << 8) | bytes[offset + 1];
    private static uint Read32(byte[] bytes, int offset) => ((uint)Read16(bytes, offset) << 16) | (uint)Read16(bytes, offset + 2);
}

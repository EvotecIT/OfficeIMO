namespace OfficeIMO.Drawing;

/// <summary>Recognizes the bounded, opaque Gray/RGB subset safe for codec pass-through.</summary>
internal static class OfficeJpeg2000Header {
    internal static bool TryGetOpaqueComponents(byte[] bytes, out int components) {
        return TryGetOpaqueDimensions(bytes, out components, out _, out _);
    }

    internal static bool TryGetOpaqueDimensions(byte[] bytes, out int components, out int width, out int height) {
        components = width = height = 0;
        if (TryReadCodestream(bytes, 0, bytes.Length, out components, out width, out height)) return true;
        if (bytes.Length < 12 || Read32(bytes, 0) != 12 || Read32(bytes, 4) != 0x6A502020 ||
            Read32(bytes, 8) != 0x0D0A870A) return false;
        bool header = false, codestream = false;
        int expected = 0, expectedWidth = 0, expectedHeight = 0;
        int offset = 12;
        while (offset < bytes.Length) {
            if (!TryReadBox(bytes, ref offset, bytes.Length, out uint type, out int start, out int end)) return false;
            if (type == 0x6A703268) { // jp2h
                if (header || !TryReadHeader(bytes, start, end, out expected, out expectedWidth, out expectedHeight)) return false;
                header = true;
            } else if (type == 0x6A703263) { // jp2c
                if (codestream || !TryReadCodestream(bytes, start, end, out components, out width, out height)) return false;
                codestream = true;
            } else if (type != 0x66747970 && type != 0x66726565 && type != 0x786D6C20 && type != 0x75756964) {
                // Extended JPX composition/channel metadata is outside this opaque subset.
                return false;
            }
        }
        return header && codestream && expected == components && expectedWidth == width && expectedHeight == height;
    }

    private static bool TryReadHeader(byte[] bytes, int offset, int end, out int components, out int width, out int height) {
        components = width = height = 0;
        int colorComponents = 0;
        while (offset < end) {
            if (!TryReadBox(bytes, ref offset, end, out uint type, out int start, out int boxEnd)) return false;
            if (type == 0x69686472) { // ihdr
                if (components != 0 || boxEnd - start != 14) return false;
                components = Read16(bytes, start + 8);
                if (!TryBoundDimensions(Read32(bytes, start + 4), Read32(bytes, start), out width, out height)) return false;
            } else if (type == 0x636F6C72) { // colr: baseline enumerated Gray or sRGB only
                if (colorComponents != 0 || boxEnd - start != 7 || bytes[start] != 1) return false;
                uint color = Read32(bytes, start + 3);
                colorComponents = color == 16 ? 3 : color == 17 ? 1 : 0;
                if (colorComponents == 0) return false;
            } else if (type == 0x63646566) { // cdef: any opacity or custom association requires normalization
                if (boxEnd - start < 2) return false;
                int count = Read16(bytes, start);
                if (boxEnd - start != 2 + count * 6) return false;
                for (int i = 0; i < count; i++) {
                    int entry = start + 2 + i * 6;
                    if (Read16(bytes, entry) != i || Read16(bytes, entry + 2) != 0 ||
                        Read16(bytes, entry + 4) != i + 1) return false;
                }
                if (count != components) return false;
            } else if (type != 0x62706363 && type != 0x72657320) { // bpcc, res
                return false; // Palette/channel remapping and extended headers are not pass-through safe.
            }
        }
        return components == colorComponents && components is 1 or 3;
    }

    private static bool TryReadCodestream(byte[] bytes, int start, int end, out int components, out int width, out int height) {
        components = width = height = 0;
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
        for (int i = 0; i < components; i++) {
            int component = start + 42 + i * 3;
            if ((bytes[component] & 127) > 15 || bytes[component + 1] == 0 || bytes[component + 2] == 0) return false;
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

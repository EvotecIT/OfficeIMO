namespace OfficeIMO.Drawing;

internal static partial class OfficeJpeg2000Decoder {
    private const uint BoxJp2Header = 0x6A703268; // jp2h
    private const uint BoxCodestream = 0x6A703263; // jp2c
    private const uint BoxImageHeader = 0x69686472; // ihdr
    private const uint BoxColour = 0x636F6C72; // colr
    private const uint BoxPalette = 0x70636C72; // pclr
    private const uint BoxComponentMapping = 0x636D6170; // cmap
    private const uint BoxChannelDefinition = 0x63646566; // cdef
    private const uint BoxBitsPerComponent = 0x62706363; // bpcc
    private const int MaximumBoxes = 4096;

    private sealed class Container {
        internal int CodestreamStart;
        internal int CodestreamEnd;
        internal ContainerHeader? Header;
    }

    /// <summary>JP2 header (I.5.3) boxes that shape the output channels.</summary>
    private sealed class ContainerHeader {
        internal OfficeJpeg2000ColorSpace ColorSpace = OfficeJpeg2000ColorSpace.Unspecified;
        internal int EnumeratedColorSpace = -1;
        internal byte[]? IccProfile;
        internal Palette? Palette;
        internal (int Component, int Type, int Column)[]? Mapping;
        internal (int Channel, int Type, int Association)[]? Definitions;
    }

    private sealed class Palette {
        internal int Entries;
        internal int[] Precisions = Array.Empty<int>();
        internal bool[] Signed = Array.Empty<bool>();
        internal int[][] Columns = Array.Empty<int[]>(); // [column][entry]
    }

    private static Container ReadContainer(byte[] bytes, CancellationToken token) {
        if (!OfficeJpeg2000Header.IsJp2Container(bytes)) {
            if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0x4F) throw new Jpeg2000Exception("The data is neither a JP2 file nor a JPEG 2000 codestream.");
            return new Container { CodestreamStart = 0, CodestreamEnd = bytes.Length };
        }
        var container = new Container { CodestreamStart = -1 };
        var header = new ContainerHeader();
        int offset = 0, boxes = 0;
        while (offset < bytes.Length) {
            token.ThrowIfCancellationRequested();
            if (++boxes > MaximumBoxes) throw new Jpeg2000Exception("The JP2 file has too many boxes.");
            if (!TryReadBox(bytes, ref offset, bytes.Length, out uint type, out int start, out int end)) break;
            if (type == BoxJp2Header) ReadHeaderBoxes(bytes, start, end, header, token);
            else if (type == BoxCodestream) {
                if (container.CodestreamStart < 0) {
                    container.CodestreamStart = start;
                    container.CodestreamEnd = end;
                }
            } else {
                // Some writers place header boxes beside jp2h; read them when jp2h did not supply them.
                ReadHeaderBox(bytes, type, start, end, header);
            }
        }
        if (container.CodestreamStart < 0) throw new Jpeg2000Exception("The JP2 file has no contiguous codestream box.");
        container.Header = header;
        return container;
    }

    private static bool TryReadBox(byte[] bytes, ref int offset, int limit, out uint type, out int start, out int end) {
        type = 0;
        start = end = 0;
        if (limit - offset < 8) return false;
        long length = ReadUInt32(bytes, offset);
        type = ReadUInt32(bytes, offset + 4);
        int headerLength = 8;
        if (length == 1) {
            if (limit - offset < 16) return false;
            length = ((long)ReadUInt32(bytes, offset + 8) << 32) | ReadUInt32(bytes, offset + 12);
            headerLength = 16;
        } else if (length == 0) {
            length = limit - offset;
        }
        if (length < headerLength) return false;
        start = offset + headerLength;
        // A truncated final box (common in embedded PDF payloads) is read to the end of the data.
        end = length > limit - offset ? limit : offset + (int)length;
        offset = end;
        return true;
    }

    private static void ReadHeaderBoxes(byte[] bytes, int offset, int limit, ContainerHeader header, CancellationToken token) {
        int boxes = 0;
        while (offset < limit) {
            token.ThrowIfCancellationRequested();
            if (++boxes > MaximumBoxes) throw new Jpeg2000Exception("The JP2 header has too many boxes.");
            if (!TryReadBox(bytes, ref offset, limit, out uint type, out int start, out int end)) break;
            ReadHeaderBox(bytes, type, start, end, header);
        }
    }

    private static void ReadHeaderBox(byte[] bytes, uint type, int start, int end, ContainerHeader header) {
        switch (type) {
            case BoxColour:
                ReadColourBox(bytes, start, end, header);
                break;
            case BoxPalette:
                if (header.Palette == null) header.Palette = ReadPalette(bytes, start, end);
                break;
            case BoxComponentMapping:
                if (header.Mapping == null) {
                    int count = (end - start) / 4;
                    var mapping = new (int, int, int)[count];
                    for (int index = 0; index < count; index++) {
                        int at = start + index * 4;
                        mapping[index] = ((bytes[at] << 8) | bytes[at + 1], bytes[at + 2], bytes[at + 3]);
                    }
                    header.Mapping = mapping;
                }
                break;
            case BoxChannelDefinition:
                if (header.Definitions == null && end - start >= 2) {
                    int count = (bytes[start] << 8) | bytes[start + 1];
                    if (2 + count * 6L > end - start) throw new Jpeg2000Exception("The JP2 channel definition box is truncated.");
                    var definitions = new (int, int, int)[count];
                    for (int index = 0; index < count; index++) {
                        int at = start + 2 + index * 6;
                        definitions[index] = ((bytes[at] << 8) | bytes[at + 1], (bytes[at + 2] << 8) | bytes[at + 3], (bytes[at + 4] << 8) | bytes[at + 5]);
                    }
                    header.Definitions = definitions;
                }
                break;
            case BoxImageHeader:
            case BoxBitsPerComponent:
                // The codestream SIZ marker is authoritative for dimensions and precisions.
                break;
        }
    }

    // Colour specification box (I.5.3.3). The first box with a usable method wins.
    private static void ReadColourBox(byte[] bytes, int start, int end, ContainerHeader header) {
        if (header.ColorSpace != OfficeJpeg2000ColorSpace.Unspecified || end - start < 3) return;
        int method = bytes[start];
        if (method == 1 && end - start >= 7) {
            int enumerated = (int)ReadUInt32(bytes, start + 3);
            header.EnumeratedColorSpace = enumerated;
            header.ColorSpace = enumerated switch {
                16 => OfficeJpeg2000ColorSpace.Srgb,
                17 => OfficeJpeg2000ColorSpace.Gray,
                18 => OfficeJpeg2000ColorSpace.Sycc,
                12 => OfficeJpeg2000ColorSpace.Cmyk,
                _ => OfficeJpeg2000ColorSpace.Other
            };
        } else if ((method == 2 || method == 3) && end - start > 3) {
            var profile = new byte[end - start - 3];
            Buffer.BlockCopy(bytes, start + 3, profile, 0, profile.Length);
            header.IccProfile = profile;
            header.ColorSpace = OfficeJpeg2000ColorSpace.Icc;
        }
    }

    // Palette box (I.5.3.4).
    private static Palette ReadPalette(byte[] bytes, int start, int end) {
        if (end - start < 3) throw new Jpeg2000Exception("The JP2 palette box is truncated.");
        int entries = (bytes[start] << 8) | bytes[start + 1];
        int columns = bytes[start + 2];
        if (entries < 1 || entries > 1024 || columns < 1) throw new Jpeg2000Exception("The JP2 palette box is invalid.");
        if (end - start < 3 + columns) throw new Jpeg2000Exception("The JP2 palette box is truncated.");
        var palette = new Palette { Entries = entries, Precisions = new int[columns], Signed = new bool[columns], Columns = new int[columns][] };
        int entryBytes = 0;
        for (int column = 0; column < columns; column++) {
            int value = bytes[start + 3 + column];
            palette.Precisions[column] = (value & 0x7F) + 1;
            palette.Signed[column] = (value & 0x80) != 0;
            if (palette.Precisions[column] > 16) throw new Jpeg2000Exception("JP2 palette precision exceeds 16 bits.");
            entryBytes += (palette.Precisions[column] + 7) / 8;
            palette.Columns[column] = new int[entries];
        }
        int position = start + 3 + columns;
        if ((long)entries * entryBytes > end - position) throw new Jpeg2000Exception("The JP2 palette box is truncated.");
        for (int entry = 0; entry < entries; entry++) {
            for (int column = 0; column < columns; column++) {
                int width = (palette.Precisions[column] + 7) / 8;
                int value = 0;
                for (int index = 0; index < width; index++) value = (value << 8) | bytes[position++];
                if (palette.Signed[column]) {
                    int shift = 32 - palette.Precisions[column];
                    value = (value << shift) >> shift;
                }
                palette.Columns[column][entry] = value;
            }
        }
        return palette;
    }

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        ((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) | ((uint)bytes[offset + 2] << 8) | bytes[offset + 3];
}

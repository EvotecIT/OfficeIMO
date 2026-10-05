using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    internal sealed class FrameHeader {
        internal int Width, Height, Left, Top, Right, Bottom;
        internal int OutputColor, BitDepth, Overlap, Transform;
        internal bool Frequency, HardTiles, IndexTable, LongWords, Trim, RedBlueNotSwapped, Premultiplied, Alpha;
        internal int[] TileWidths = Array.Empty<int>(), TileHeights = Array.Empty<int>();
        internal PlaneHeader Primary = new();
        internal PlaneHeader? AlphaPlane;
        internal int HeaderEnd;
    }
    internal sealed class PlaneHeader {
        internal int Color, Components, Bands, ShiftBits;
        internal bool Scaled;
        internal int[]? DcQuant, LpQuant, HpQuant;
    }

    // T.832 8.3/8.4. This parses the bounded unsigned RGB/gray frame contract; pixel
    // reconstruction and container dispatch remain separate codec responsibilities.
    internal static FrameHeader ReadHeader(byte[] bytes, int offset, int length, CancellationToken token) {
        var bits = new Bits(bytes, offset, length, token);
        if (bits.Read(32) != 0x574D5048 || bits.Read(32) != 0x4F544F00 || bits.Read(4) != 1)
            throw new FormatException("JPEG-XR image signature or version is invalid.");
        var frame = new FrameHeader { HardTiles = bits.Flag() };
        bits.Read(3); // RESERVED_C must be ignored by conforming decoders.
        bool tiled = bits.Flag();
        frame.Frequency = bits.Flag(); frame.Transform = (int)bits.Read(3);
        frame.IndexTable = bits.Flag(); frame.Overlap = (int)bits.Read(2);
        bool shortHeader = bits.Flag(); frame.LongWords = bits.Flag();
        bool window = bits.Flag(); frame.Trim = bits.Flag(); bits.Read(1);
        frame.RedBlueNotSwapped = bits.Flag(); frame.Premultiplied = bits.Flag(); frame.Alpha = bits.Flag();
        frame.OutputColor = (int)bits.Read(4); frame.BitDepth = (int)bits.Read(4);
        uint width = bits.Read(shortHeader ? 16 : 32), height = bits.Read(shortHeader ? 16 : 32);
        if (width >= int.MaxValue || height >= int.MaxValue || frame.Overlap == 3 ||
            (frame.BitDepth != 1 && frame.BitDepth != 2) || (frame.OutputColor != 0 && frame.OutputColor != 7))
            throw new FormatException("JPEG-XR frame is outside the unsigned eight/sixteen-bit RGB/gray contract.");
        frame.Width = (int)width + 1; frame.Height = (int)height + 1;
        if ((long)frame.Width * frame.Height > 50_000_000L)
            throw new FormatException("JPEG-XR image dimensions exceed the managed limit.");
        int columns = tiled ? (int)bits.Read(12) + 1 : 1;
        int rows = tiled ? (int)bits.Read(12) + 1 : 1;
        frame.TileWidths = new int[columns]; frame.TileHeights = new int[rows];
        int widthSum = 0, heightSum = 0;
        for (int i = 0; i < columns - 1; i++) {
            token.ThrowIfCancellationRequested();
            frame.TileWidths[i] = (int)bits.Read(shortHeader ? 8 : 16); widthSum += frame.TileWidths[i];
            if (frame.TileWidths[i] == 0) throw new FormatException("JPEG-XR tile width is zero.");
        }
        for (int i = 0; i < rows - 1; i++) {
            token.ThrowIfCancellationRequested();
            frame.TileHeights[i] = (int)bits.Read(shortHeader ? 8 : 16); heightSum += frame.TileHeights[i];
            if (frame.TileHeights[i] == 0) throw new FormatException("JPEG-XR tile height is zero.");
        }
        if (window) {
            frame.Top = (int)bits.Read(6); frame.Left = (int)bits.Read(6);
            frame.Bottom = (int)bits.Read(6); frame.Right = (int)bits.Read(6);
        } else {
            frame.Right = (16 - frame.Width % 16) % 16;
            frame.Bottom = (16 - frame.Height % 16) % 16;
        }
        int paddedWidth = checked(frame.Width + frame.Left + frame.Right);
        int paddedHeight = checked(frame.Height + frame.Top + frame.Bottom);
        if (paddedWidth % 16 != 0 || paddedHeight % 16 != 0 ||
            columns > paddedWidth / 16 || rows > paddedHeight / 16 ||
            widthSum >= paddedWidth / 16 || heightSum >= paddedHeight / 16)
            throw new FormatException("JPEG-XR tile/window geometry is invalid.");
        frame.TileWidths[columns - 1] = paddedWidth / 16 - widthSum;
        frame.TileHeights[rows - 1] = paddedHeight / 16 - heightSum;
        frame.Primary = ReadPlaneHeader(bits, false, frame.BitDepth);
        if (frame.Alpha) frame.AlphaPlane = ReadPlaneHeader(bits, true, frame.BitDepth);
        if (frame.AlphaPlane != null && frame.AlphaPlane.Bands != frame.Primary.Bands)
            throw new FormatException("JPEG-XR differing interleaved alpha subbands are outside the managed contract.");
        frame.HeaderEnd = bits.ByteOffset;
        return frame;
    }

    private static PlaneHeader ReadPlaneHeader(Bits bits, bool alpha, int bitDepth) {
        var plane = new PlaneHeader { Color = (int)bits.Read(3), Scaled = bits.Flag(), Bands = (int)bits.Read(4) };
        if ((plane.Color != 0 && plane.Color != 3) || alpha && plane.Color != 0 || plane.Bands > 3)
            throw new FormatException("JPEG-XR plane is outside the gray/YUV444 contract.");
        plane.Components = plane.Color == 0 ? 1 : 3;
        if (plane.Color == 3) bits.Read(8); // Reserved fields are ignored.
        if (bitDepth == 2) plane.ShiftBits = (int)bits.Read(8);
        if (bits.Flag()) plane.DcQuant = ReadQuantization(bits, plane.Components);
        if (plane.Bands != 3) {
            bits.Read(1);
            if (bits.Flag()) plane.LpQuant = ReadQuantization(bits, plane.Components);
            if (plane.Bands < 2) {
                bits.Read(1);
                if (bits.Flag()) plane.HpQuant = ReadQuantization(bits, plane.Components);
            }
        }
        bits.AlignZero();
        return plane;
    }

    private static int[] ReadQuantization(Bits bits, int components) {
        int mode = components == 1 ? 0 : (int)bits.Read(2);
        if (mode == 3) throw new FormatException("JPEG-XR quantization component mode is reserved.");
        var values = new int[components]; values[0] = (int)bits.Read(8);
        if (mode == 0) { for (int i = 1; i < components; i++) values[i] = values[0]; }
        else if (mode == 1) { int chroma = (int)bits.Read(8); for (int i = 1; i < components; i++) values[i] = chroma; }
        else { for (int i = 1; i < components; i++) values[i] = (int)bits.Read(8); }
        return values;
    }
}

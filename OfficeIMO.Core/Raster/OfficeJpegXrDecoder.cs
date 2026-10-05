using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    internal static bool HasSignature(byte[] bytes) => bytes.Length >= 4 &&
        bytes[0] == 0x49 && bytes[1] == 0x49 && bytes[2] == 0xBC && bytes[3] == 1;

    internal static bool TryIdentify(byte[] bytes, CancellationToken cancellation, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (!HasSignature(bytes)) return false;
        try {
            Container container = ReadContainer(bytes, cancellation);
            bool swap = container.Transform >= 4;
            info = new OfficeImageInfo(OfficeImageFormat.JpegXr,
                swap ? container.Height : container.Width, swap ? container.Width : container.Height,
                swap ? container.DpiY : container.DpiX, swap ? container.DpiX : container.DpiY);
            return true;
        } catch (FormatException) { return false; }
        catch (OverflowException) { return false; }
    }

    internal static bool TryDecode(byte[] bytes, OfficeRasterDecodeOptions options, out OfficeRasterImage? image, OfficeIccColorProfile? colorProfile = null) {
        image = null;
        options.CancellationToken.ThrowIfCancellationRequested();
        if (bytes.Length > options.MaximumEncodedBytes || options.FrameIndex != 0) return false;
        try {
            Container container = ReadContainer(bytes, options.CancellationToken);
            if (colorProfile != null && colorProfile.ComponentCount != (container.Gray ? 1 : 3)) return false;
            if (!OfficeRasterGuards.TryEnsurePixelCount(container.Width, container.Height, options.MaximumDecodedPixels, out int pixels)) return false;
            // Budget the complete pipeline before allocating coefficient/sample arrays.
            // This includes simultaneously retained alpha, prediction scratch, tile
            // quantizer tables, output, and a possible oriented output copy.
            long workingBytes = checked(bytes.LongLength + options.RetainedManagedBytes + pixels * 8L + 1024 * 1024);
            workingBytes = checked(workingBytes + FrameWorkingBytes(container.Frame));
            if (container.SeparateAlpha != null) workingBytes = checked(workingBytes + FrameWorkingBytes(container.SeparateAlpha));
            if (workingBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
            int[][] primary = DecodeFrame(bytes, container.ImageOffset, container.ImageLength, container.Frame,
                options.CancellationToken, out int[][]? alpha);
            FrameHeader? alphaFrame = container.Frame.Alpha ? container.Frame : null;
            if (container.SeparateAlpha != null) {
                alphaFrame = container.SeparateAlpha;
                alpha = DecodeFrame(bytes, container.AlphaOffset, container.AlphaLength, alphaFrame, options.CancellationToken, out _);
            }
            byte[] rgba = FormatRgba(container.Frame, primary, alphaFrame, alpha, container.Premultiplied, colorProfile, options.CancellationToken);
            int width = container.Width, height = container.Height;
            int orientation = container.Transform switch { 1 => 4, 2 => 2, 3 => 3, 4 => 6, 5 => 7, 6 => 5, 7 => 8, _ => 1 };
            rgba = OfficeRasterOrientation.Apply(rgba, ref width, ref height, orientation, options.CancellationToken,
                "JPEG-XR orientation exceeds the raster limit.");
            image = OfficeRasterImage.FromOwnedRgba32(width, height, rgba);
            return true;
        } catch (FormatException) { return false; }
        catch (OverflowException) { return false; }
    }

    private static long FrameWorkingBytes(FrameHeader frame) {
        long pixels = checked((long)(frame.Width + frame.Left + frame.Right) * (frame.Height + frame.Top + frame.Bottom));
        long components = frame.Primary.Components + (frame.AlphaPlane?.Components ?? 0);
        long tiles = (long)frame.TileWidths.Length * frame.TileHeights.Length;
        // Twelve bytes per padded component sample cover HP and sample arrays,
        // all DC/LP prediction copies, per-macroblock contexts, and flexbits state.
        return checked(pixels * components * 12 + tiles * 8192);
    }

    private static int[][] DecodeFrame(byte[] bytes, int offset, int length, FrameHeader frame,
            CancellationToken cancellation, out int[][]? alpha) {
        PacketMap packets = ReadPackets(bytes, offset, length, frame, cancellation);
        ReconstructedFrame reconstructed;
        if (frame.Frequency) {
            BandData dc = ReadFrequencyDc(bytes, frame, packets, cancellation);
            BandData? lp = frame.Primary.Bands < 3 ? ReadFrequencyLp(bytes, frame, packets, dc, cancellation) : null;
            DclpPlane primary = ReconstructDclp(frame, dc, lp, false, cancellation);
            DclpPlane? alphaPlane = frame.Alpha ? ReconstructDclp(frame, dc, lp, true, cancellation) : null;
            BandData? hp = frame.Primary.Bands < 2 ? ReadFrequencyHp(bytes, frame, packets, lp!, primary, alphaPlane, cancellation) : null;
            if (hp != null) {
                if (frame.Primary.Bands == 0) ReadFrequencyFlex(bytes, frame, packets, hp, cancellation);
                ReconstructHighpass(frame, hp, primary, alphaPlane, cancellation);
            }
            reconstructed = new ReconstructedFrame { Primary = primary, Alpha = alphaPlane, Highpass = hp };
        } else reconstructed = ReadSpatial(bytes, frame, packets, cancellation);
        int[][] samples = ReconstructSamples(frame, frame.Primary, reconstructed.Primary, reconstructed.Highpass?.Primary, cancellation);
        alpha = frame.AlphaPlane == null ? null : ReconstructSamples(frame, frame.AlphaPlane, reconstructed.Alpha!, reconstructed.Highpass?.Alpha, cancellation);
        return samples;
    }
}

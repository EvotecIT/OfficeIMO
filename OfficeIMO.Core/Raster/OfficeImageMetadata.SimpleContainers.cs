using System;
using System.Threading;
using OfficeIMO.Provenance;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeImageMetadata {
    private const uint BitmapProfileLinked = 0x4C494E4B;
    private const uint BitmapProfileEmbedded = 0x4D424544;

    private static byte[] CloneWithoutProfiles(byte[] input, OfficeImageFormat format, CancellationToken token, out OfficeImageMetadataProfileKinds present) {
        if (!OfficeImageReader.TryValidateContent(input, null, token, out OfficeImageInfo info) || info.Format != format) throw new FormatException("The image payload is malformed.");
        present = OfficeImageMetadataProfileKinds.None;
        return (byte[])input.Clone();
    }

    private static void ReadGifProfiles(byte[] input, OfficeImageMetadata metadata, CancellationToken token) {
        _ = OfficeProvenanceGif.RewriteMetadataProfiles(input, null, null, OfficeImageMetadataProfileKinds.None,
            false, token, out _, out byte[]? xmp, out byte[]? icc, readOnly: true);
        metadata.XmpProfile = xmp;
        metadata.IccProfile = icc;
        metadata.ResolutionUnits = OfficeImageResolutionUnit.AspectRatio;
        metadata.HorizontalResolution = input[12] == 0 ? 1D : (input[12] + 15D) / 64D;
        metadata.VerticalResolution = 1D;
    }

    private static byte[] RewriteGif(byte[] input, OfficeImageMetadata? metadata, OfficeImageMetadataProfileKinds replace,
        CancellationToken token, out OfficeImageMetadataProfileKinds present, long additionallyRetainedBytes = 0L) {
        if (metadata != null && (metadata.HasExifProfile || metadata._iptc != null)) throw new NotSupportedException("GIF application metadata supports XMP and ICC profiles; Exif and IPTC IIM profiles have no standard carrier.");
        byte[] output = OfficeProvenanceGif.RewriteMetadataProfiles(input, metadata?._xmp, metadata?._icc,
            replace, metadata != null, token, out present, out _, out _, additionallyRetainedBytes: additionallyRetainedBytes);
        if (metadata != null) {
            if (metadata.ResolutionUnits != OfficeImageResolutionUnit.AspectRatio) throw new NotSupportedException("GIF has no physical-density carrier. Use an aspect-ratio resolution or PrepareForEncoding for GIF output.");
            double ratio = metadata.HorizontalResolution / metadata.VerticalResolution;
            if (ratio != 1D || input[12] != 0) {
                double value = Math.Round(ratio * 64D - 15D, MidpointRounding.AwayFromZero);
                if (value < 1D || value > 255D) throw new ArgumentOutOfRangeException(nameof(metadata), "The aspect ratio is outside GIF's representable range.");
                output[12] = (byte)value;
            }
        }
        return output;
    }

    private static void ReadBmpProfiles(byte[] input, OfficeImageMetadata metadata) {
        if (input.Length < 54) return;
        int header = checked((int)OfficeExifProfileCodec.Read(input, 14, 4, true));
        if (header >= 40) {
            int x = unchecked((int)OfficeExifProfileCodec.Read(input, 38, 4, true));
            int y = unchecked((int)OfficeExifProfileCodec.Read(input, 42, 4, true));
            if (x > 0 && y > 0) {
                metadata.HorizontalResolution = x;
                metadata.VerticalResolution = y;
                metadata.ResolutionUnits = OfficeImageResolutionUnit.PixelsPerMeter;
            }
        }
        if (TryGetBmpProfile(input, out int offset, out int length, out bool embedded) && embedded) metadata.IccProfile = Slice(input, offset, length);
    }

    private static byte[] RewriteBmp(byte[] input, OfficeImageMetadata? metadata, OfficeImageMetadataProfileKinds replace, out OfficeImageMetadataProfileKinds present, long additionallyRetainedBytes = 0L) {
        present = OfficeImageMetadataProfileKinds.None;
        if (metadata != null && (metadata.HasExifProfile || metadata._xmp != null || metadata._iptc != null)) throw new NotSupportedException("BMP supports ICC color profiles and density; Exif, XMP, and IPTC IIM profiles have no standard BMP carrier.");
        if (checked(input.LongLength * (metadata?._icc != null ? 3L : 2L) + additionallyRetainedBytes + (metadata?.RetainedProfileBytes ?? 0L) * 2L + 256L) > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("Metadata rewriting exceeds the managed working-set limit.");
        byte[] output = (byte[])input.Clone();
        int header = checked((int)OfficeExifProfileCodec.Read(input, 14, 4, true));
        if (TryGetBmpProfile(input, out int profileOffset, out int profileLength, out _)) {
            present = OfficeImageMetadataProfileKinds.Icc;
            if ((replace & OfficeImageMetadataProfileKinds.Icc) != 0) {
                Array.Clear(output, profileOffset, profileLength);
                OfficeExifProfileCodec.Write(output, 14 + 56, 0x73524742, 4, true);
                OfficeExifProfileCodec.Write(output, 14 + 112, 0, 4, true);
                OfficeExifProfileCodec.Write(output, 14 + 116, 0, 4, true);
            }
        }
        if (metadata == null) return output;
        if (header < 40 || header > 124) throw new NotSupportedException("BMP metadata replacement requires a Windows bitmap header.");
        double scale = metadata.ResolutionUnits == OfficeImageResolutionUnit.PixelsPerInch ? 1D / 0.0254D : metadata.ResolutionUnits == OfficeImageResolutionUnit.PixelsPerCentimeter ? 100D : 1D;
        if (metadata.ResolutionUnits == OfficeImageResolutionUnit.AspectRatio) {
            OfficeExifProfileCodec.Write(output, 38, 0, 4, true);
            OfficeExifProfileCodec.Write(output, 42, 0, 4, true);
        } else {
            OfficeExifProfileCodec.Write(output, 38, Density(metadata.HorizontalResolution * scale, int.MaxValue), 4, true);
            OfficeExifProfileCodec.Write(output, 42, Density(metadata.VerticalResolution * scale, int.MaxValue), 4, true);
        }
        if (metadata._icc == null) return output;
        if (!OfficeIccProfileValidator.TryValidate(metadata._icc, 0, metadata._icc.Length)) throw new FormatException("The ICC profile is malformed.");
        int growth = 124 - header;
        int newLength = checked(output.Length + growth + metadata._icc.Length);
        if (!OfficeRasterGuards.IsEncodedPayloadWithinLimits(newLength)) throw new FormatException("The edited BMP exceeds the encoded-size limit.");
        var expanded = new byte[newLength];
        Buffer.BlockCopy(output, 0, expanded, 0, 14 + header);
        Buffer.BlockCopy(output, 14 + header, expanded, 138, output.Length - 14 - header);
        // BITMAPINFOHEADER stores bit masks after the header; V5 stores them inside it.
        uint compression = (uint)OfficeExifProfileCodec.Read(output, 30, 4, true);
        if (header == 40 && (compression == 3 || compression == 6)) {
            int maskBytes = compression == 6 ? 16 : 12;
            if (14 + header + maskBytes > output.Length) throw new FormatException("Truncated BMP channel masks.");
            Buffer.BlockCopy(output, 54, expanded, 54, maskBytes);
        }
        OfficeExifProfileCodec.Write(expanded, 2, (uint)newLength, 4, true);
        OfficeExifProfileCodec.Write(expanded, 10, OfficeExifProfileCodec.Read(output, 10, 4, true) + (uint)growth, 4, true);
        OfficeExifProfileCodec.Write(expanded, 14, 124, 4, true);
        OfficeExifProfileCodec.Write(expanded, 70, BitmapProfileEmbedded, 4, true);
        OfficeExifProfileCodec.Write(expanded, 126, (uint)(output.Length + growth - 14), 4, true);
        OfficeExifProfileCodec.Write(expanded, 130, (uint)metadata._icc.Length, 4, true);
        Buffer.BlockCopy(metadata._icc, 0, expanded, output.Length + growth, metadata._icc.Length);
        return expanded;
    }

    private static bool TryGetBmpProfile(byte[] input, out int offset, out int length, out bool embedded) {
        offset = 0;
        length = 0;
        embedded = false;
        if (input.Length < 138 || OfficeExifProfileCodec.Read(input, 14, 4, true) != 124) return false;
        uint kind = (uint)OfficeExifProfileCodec.Read(input, 70, 4, true);
        if (kind != BitmapProfileEmbedded && kind != BitmapProfileLinked) return false;
        long start = 14L + (long)OfficeExifProfileCodec.Read(input, 126, 4, true);
        long size = (long)OfficeExifProfileCodec.Read(input, 130, 4, true);
        long pixels = (long)OfficeExifProfileCodec.Read(input, 10, 4, true);
        long pixelBytes = (long)OfficeExifProfileCodec.Read(input, 34, 4, true);
        if (pixelBytes == 0) {
            long width = (long)OfficeExifProfileCodec.Read(input, 18, 4, true);
            int height = unchecked((int)OfficeExifProfileCodec.Read(input, 22, 4, true));
            long bits = (long)OfficeExifProfileCodec.Read(input, 28, 2, true);
            pixelBytes = checked(((width * bits + 31) / 32) * 4 * Math.Abs((long)height));
        }
        if (start < 138 || size <= 0 || size > OfficeExifProfileCodec.MaximumProfileBytes || start + size > input.Length || start < pixels + pixelBytes && start + size > pixels) throw new FormatException("BMP color-profile range is invalid or overlaps pixels.");
        offset = checked((int)start);
        length = checked((int)size);
        embedded = kind == BitmapProfileEmbedded;
        return true;
    }
}

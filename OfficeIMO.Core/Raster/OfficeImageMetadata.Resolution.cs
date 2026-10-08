using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeImageMetadata {
    private static byte[]? EncodeDensityExif(OfficeImageMetadata metadata, System.Threading.CancellationToken token) {
        OfficeImageMetadata copy = metadata.Clone();
        GetExifResolution(metadata, out double x, out double y, out ushort unit);
        copy._changes[OfficeExifTag.XResolution] = GetResolutionValue(OfficeExifTag.XResolution, x, metadata.GetExifValue(OfficeExifTag.XResolution));
        copy._changes[OfficeExifTag.YResolution] = GetResolutionValue(OfficeExifTag.YResolution, y, metadata.GetExifValue(OfficeExifTag.YResolution));
        copy._removed.Remove(OfficeExifTag.XResolution); copy._removed.Remove(OfficeExifTag.YResolution);
        copy.SetExifValue(OfficeExifTag.ResolutionUnit, unit);
        return copy.EncodeExifProfile(token);
    }

    private static OfficeExifValue GetResolutionValue(OfficeExifTag tag, double density, OfficeExifValue? existing) {
        if (existing?.Value is OfficeRational rational && rational.Numerator != 0 && rational.Denominator != 0 && rational.ToDouble() == density) return existing;
        return new OfficeExifValue(tag, OfficeUnsignedRational.FromPositiveDouble(density));
    }

    private static bool HasJpegJfif(byte[] input) {
        for (int cursor = 2; cursor < input.Length && OfficeIMO.Provenance.OfficeProvenanceJpeg.TryReadMarker(input, cursor, out byte marker, out int payload, out int length, out int end); cursor = end) {
            if (marker == 0xDA || marker == 0xD9) break;
            if (marker == 0xE0 && length >= 12 && PrefixAt(input, payload, length, System.Text.Encoding.ASCII.GetBytes("JFIF\0"))) return true;
        }
        return false;
    }
    private static void GetExifResolution(OfficeImageMetadata metadata, out double x, out double y, out ushort unit) {
        x = metadata.HorizontalResolution;
        y = metadata.VerticalResolution;
        unit = metadata.ResolutionUnits == OfficeImageResolutionUnit.AspectRatio ? (ushort)1 : metadata.ResolutionUnits == OfficeImageResolutionUnit.PixelsPerCentimeter || metadata.ResolutionUnits == OfficeImageResolutionUnit.PixelsPerMeter ? (ushort)3 : (ushort)2;
        if (metadata.ResolutionUnits == OfficeImageResolutionUnit.PixelsPerMeter) { x /= 100D; y /= 100D; }
    }

    private static void ReadNativeResolution(byte[] input, OfficeImageFormat format, OfficeImageMetadata metadata, System.Threading.CancellationToken token) {
        if (format == OfficeImageFormat.Png) {
            for (int cursor = 8; cursor <= input.Length - 12;) {
                token.ThrowIfCancellationRequested();
                int length = checked((int)OfficeExifProfileCodec.Read(input, cursor, 4, false));
                if (length > input.Length - cursor - 12) throw new FormatException("Truncated PNG density chunk.");
                if (PrefixAt(input, cursor + 4, 4, System.Text.Encoding.ASCII.GetBytes("pHYs")) && length == 9) {
                    uint x = (uint)OfficeExifProfileCodec.Read(input, cursor + 8, 4, false);
                    uint y = (uint)OfficeExifProfileCodec.Read(input, cursor + 12, 4, false);
                    if (x > 0 && y > 0) { metadata.HorizontalResolution = x; metadata.VerticalResolution = y; metadata.ResolutionUnits = input[cursor + 16] == 0 ? OfficeImageResolutionUnit.AspectRatio : OfficeImageResolutionUnit.PixelsPerMeter; }
                }
                cursor += length + 12;
            }
        } else if (format == OfficeImageFormat.Jpeg) {
            for (int cursor = 2; cursor < input.Length && OfficeIMO.Provenance.OfficeProvenanceJpeg.TryReadMarker(input, cursor, out byte marker, out int payload, out int length, out int end); cursor = end) {
                token.ThrowIfCancellationRequested();
                if (marker == 0xDA || marker == 0xD9) break;
                if (marker != 0xE0 || length < 12 || !PrefixAt(input, payload, length, System.Text.Encoding.ASCII.GetBytes("JFIF\0"))) continue;
                uint x = (uint)OfficeExifProfileCodec.Read(input, payload + 8, 2, false);
                uint y = (uint)OfficeExifProfileCodec.Read(input, payload + 10, 2, false);
                if (x > 0 && y > 0) { metadata.HorizontalResolution = x; metadata.VerticalResolution = y; metadata.ResolutionUnits = input[payload + 7] == 0 ? OfficeImageResolutionUnit.AspectRatio : input[payload + 7] == 2 ? OfficeImageResolutionUnit.PixelsPerCentimeter : OfficeImageResolutionUnit.PixelsPerInch; }
                return;
            }
            ReadExifResolution(metadata);
        } else if (format == OfficeImageFormat.Webp) {
            ReadExifResolution(metadata);
        }
    }

    private static void ReadExifResolution(OfficeImageMetadata metadata) {
            OfficeExifValue? xv = metadata.GetExifValue(OfficeExifTag.XResolution);
            OfficeExifValue? yv = metadata.GetExifValue(OfficeExifTag.YResolution);
            OfficeExifValue? uv = metadata.GetExifValue(OfficeExifTag.ResolutionUnit);
            if (xv?.Value is OfficeRational x && yv?.Value is OfficeRational y && x.ToDouble() > 0 && y.ToDouble() > 0) {
                ushort unit = uv?.Value is ushort value ? value : (ushort)2;
                metadata.HorizontalResolution = x.ToDouble(); metadata.VerticalResolution = y.ToDouble();
                metadata.ResolutionUnits = unit == 1 ? OfficeImageResolutionUnit.AspectRatio : unit == 3 ? OfficeImageResolutionUnit.PixelsPerCentimeter : OfficeImageResolutionUnit.PixelsPerInch;
            }
    }
}

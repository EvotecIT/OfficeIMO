using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeImageMetadata {
    private static byte[]? EncodeDensityExif(OfficeImageMetadata metadata, System.Threading.CancellationToken token) {
        OfficeImageMetadata copy = metadata.Clone();
        GetExifResolution(metadata, out double x, out double y, out ushort unit);
        if (!copy._removed.Contains(OfficeExifTag.XResolution)) copy._changes[OfficeExifTag.XResolution] = GetResolutionValue(OfficeExifTag.XResolution, x, metadata.GetExifValue(OfficeExifTag.XResolution));
        if (!copy._removed.Contains(OfficeExifTag.YResolution)) copy._changes[OfficeExifTag.YResolution] = GetResolutionValue(OfficeExifTag.YResolution, y, metadata.GetExifValue(OfficeExifTag.YResolution));
        if (!copy._removed.Contains(OfficeExifTag.ResolutionUnit)) copy._changes[OfficeExifTag.ResolutionUnit] = new OfficeExifValue(OfficeExifTag.ResolutionUnit, unit);
        return copy.EncodeExifProfile(token);
    }

    private static OfficeExifValue GetResolutionValue(OfficeExifTag tag, double density, OfficeExifValue? existing) {
        if (existing?.Value is OfficeRational rational && rational.Numerator != 0 && rational.Denominator != 0 && rational.ToDouble() == density) return existing;
        return new OfficeExifValue(tag, OfficeUnsignedRational.FromPositiveDouble(density));
    }

    private static bool TryFindJpegJfif(byte[] input, System.Threading.CancellationToken token, out int start, out int payloadOffset, out int payloadLength) {
        for (int cursor = 2; cursor < input.Length && OfficeIMO.Provenance.OfficeProvenanceJpeg.TryReadMarker(input, cursor, out byte marker, out int payload, out int length, out int end); cursor = end) {
            token.ThrowIfCancellationRequested();
            if (marker == 0xDA || marker == 0xD9) break;
            if (marker == 0xE0 && length >= 12 && PrefixAt(input, payload, length, JpegJfifPrefix)) { start = cursor; payloadOffset = payload; payloadLength = length; return true; }
        }
        start = payloadOffset = payloadLength = -1;
        return false;
    }
    private static void GetIntegerResolution(double x, double y, bool aspectRatio, uint maximum, out uint horizontal, out uint vertical) {
        if (!aspectRatio || x >= 1D && y >= 1D && x <= maximum && y <= maximum && x == Math.Truncate(x) && y == Math.Truncate(y)) {
            horizontal = Density(x, maximum); vertical = Density(y, maximum);
            return;
        }
        // Unitless carriers store a pair of integer words, so preserve the ratio
        // rather than rounding each fractional component independently.
        OfficeRational ratio = OfficeUnsignedRational.FromPositiveDouble(x / y, "resolution", maximum);
        horizontal = ratio.Numerator; vertical = ratio.Denominator;
    }
    private static void GetExifResolution(OfficeImageMetadata metadata, out double x, out double y, out ushort unit) => GetExifResolution(metadata.Resolution, out x, out y, out unit);
    private static void GetExifResolution(OfficeImageResolution resolution, out double x, out double y, out ushort unit) {
        x = resolution.Horizontal;
        y = resolution.Vertical;
        unit = resolution.Unit == OfficeImageResolutionUnit.AspectRatio ? (ushort)1 : resolution.Unit == OfficeImageResolutionUnit.PixelsPerCentimeter || resolution.Unit == OfficeImageResolutionUnit.PixelsPerMeter ? (ushort)3 : (ushort)2;
        if (resolution.Unit == OfficeImageResolutionUnit.PixelsPerMeter) { x /= 100D; y /= 100D; }
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
                    if (x > 0 && y > 0) metadata._resolution = new OfficeImageResolution(x, y, input[cursor + 16] == 0 ? OfficeImageResolutionUnit.AspectRatio : OfficeImageResolutionUnit.PixelsPerMeter);
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
                if (x > 0 && y > 0) metadata._resolution = new OfficeImageResolution(x, y, input[payload + 7] == 0 ? OfficeImageResolutionUnit.AspectRatio : input[payload + 7] == 2 ? OfficeImageResolutionUnit.PixelsPerCentimeter : OfficeImageResolutionUnit.PixelsPerInch);
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
            if (xv?.Value is OfficeRational || yv?.Value is OfficeRational || uv?.Value is ushort) {
                ushort unit = uv?.Value is ushort value ? value : (ushort)2;
                double x = xv?.Value is OfficeRational xr && xr.ToDouble() > 0 ? xr.ToDouble() : metadata.HorizontalResolution;
                double y = yv?.Value is OfficeRational yr && yr.ToDouble() > 0 ? yr.ToDouble() : metadata.VerticalResolution;
                metadata._resolution = new OfficeImageResolution(x, y, unit == 1 ? OfficeImageResolutionUnit.AspectRatio : unit == 3 ? OfficeImageResolutionUnit.PixelsPerCentimeter : OfficeImageResolutionUnit.PixelsPerInch);
            }
    }

    private OfficeImageResolution ResolveExifDensityEdit(OfficeExifTag tag, object value) {
        ushort currentUnit = GetExifValue(OfficeExifTag.ResolutionUnit)?.Value is ushort storedUnit ? storedUnit : (ushort)2;
        double x = _resolution.Horizontal, y = _resolution.Vertical;
        if (currentUnit != 1 && _resolution.PhysicalDpiX.HasValue) {
            double unitScale = currentUnit == 3 ? 2.54D : 1D;
            x = _resolution.PhysicalDpiX.Value / unitScale;
            y = _resolution.PhysicalDpiY!.Value / unitScale;
        }
        OfficeImageResolutionUnit nativeUnit = currentUnit == 1 ? OfficeImageResolutionUnit.AspectRatio : currentUnit == 3 ? OfficeImageResolutionUnit.PixelsPerCentimeter : OfficeImageResolutionUnit.PixelsPerInch;
        if (tag.Equals(OfficeExifTag.XResolution) || tag.Equals(OfficeExifTag.YResolution)) {
            if (!(value is OfficeRational rational)) throw new ArgumentException("Exif density requires an unsigned rational value.", nameof(value));
            double density = rational.ToDouble();
            return new OfficeImageResolution(tag.Equals(OfficeExifTag.XResolution) ? density : x,
                tag.Equals(OfficeExifTag.YResolution) ? density : y, nativeUnit);
        }
        if (tag.Equals(OfficeExifTag.ResolutionUnit)) {
            if (!(value is ushort unit) || unit < 1 || unit > 3) throw new ArgumentException("Exif resolution unit must be 1 (ratio), 2 (inches), or 3 (centimeters).", nameof(value));
            // Exif has no meter carrier. The authored unit changes the interpretation
            // of the existing native values, just as the typed Resolution assignment does.
            return new OfficeImageResolution(x, y,
                unit == 1 ? OfficeImageResolutionUnit.AspectRatio : unit == 3 ? OfficeImageResolutionUnit.PixelsPerCentimeter : OfficeImageResolutionUnit.PixelsPerInch);
        }
        return _resolution;
    }

    private void SetResolutionAndDensityFields(OfficeImageResolution resolution, bool restoreRemoved = true, OfficeExifValue? edited = null) {
        GetExifResolution(resolution, out double x, out double y, out ushort unit);
        OfficeExifValue? existingX = edited?.Tag.Equals(OfficeExifTag.XResolution) == true ? edited : GetExifValue(OfficeExifTag.XResolution);
        OfficeExifValue? existingY = edited?.Tag.Equals(OfficeExifTag.YResolution) == true ? edited : GetExifValue(OfficeExifTag.YResolution);
        OfficeExifValue? existingUnit = edited?.Tag.Equals(OfficeExifTag.ResolutionUnit) == true ? edited : GetExifValue(OfficeExifTag.ResolutionUnit);
        bool densityPresent = existingX != null || existingY != null || existingUnit != null;
        bool restoreDensity = restoreRemoved && (densityPresent || _removed.Contains(OfficeExifTag.XResolution) || _removed.Contains(OfficeExifTag.YResolution) || _removed.Contains(OfficeExifTag.ResolutionUnit));
        // Validate every replacement before publishing any part of the edit. In particular,
        // a valid X replacement must not survive rejection of an unrepresentable Y value.
        OfficeExifValue? changedX = existingX != null || restoreDensity ? GetResolutionValue(OfficeExifTag.XResolution, x, existingX) : null;
        OfficeExifValue? changedY = existingY != null || restoreDensity ? GetResolutionValue(OfficeExifTag.YResolution, y, existingY) : null;
        OfficeExifValue? changedUnit = existingUnit != null || restoreDensity ? new OfficeExifValue(OfficeExifTag.ResolutionUnit, unit) : null;
        if (changedX != null) _changes[OfficeExifTag.XResolution] = changedX;
        if (changedY != null) _changes[OfficeExifTag.YResolution] = changedY;
        if (changedUnit != null) _changes[OfficeExifTag.ResolutionUnit] = changedUnit;
        _resolution = resolution;
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OfficeIMO.Provenance;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeImageMetadata {
    private static OfficeImageResolutionEvidence InspectResolution(OfficeImageFormat format, byte[]? bytes,
        List<OfficeExifFieldEvidence> fields, List<string> issues, CancellationToken token) {
        var candidates = new List<OfficeImageResolutionEvidence>();
        if (bytes != null && format == OfficeImageFormat.Jpeg) {
            for (int cursor = 2; cursor < bytes.Length && OfficeProvenanceJpeg.TryReadMarker(bytes, cursor, out byte marker, out int payload, out int length, out int end); cursor = end) {
                token.ThrowIfCancellationRequested();
                if (marker == 0xDA || marker == 0xD9) break;
                if (marker != 0xE0 || !PrefixAt(bytes, payload, length, JpegJfifPrefix)) continue;
                if (length < 12) { candidates.Add(new OfficeImageResolutionEvidence(true, "JPEG/JFIF", null, false)); continue; }
                int unit = bytes[payload + 7];
                double x = OfficeExifProfileCodec.Read(bytes, payload + 8, 2, false), y = OfficeExifProfileCodec.Read(bytes, payload + 10, 2, false);
                Add("JPEG/JFIF", x, y, unit == 0 ? OfficeImageResolutionUnit.AspectRatio : unit == 2 ? OfficeImageResolutionUnit.PixelsPerCentimeter : OfficeImageResolutionUnit.PixelsPerInch, unit <= 2);
            }
        }
        if (bytes != null && format == OfficeImageFormat.Png) {
            for (int cursor = 8; cursor <= bytes.Length - 12;) {
                token.ThrowIfCancellationRequested();
                ulong length = OfficeExifProfileCodec.Read(bytes, cursor, 4, false);
                if (length > (ulong)(bytes.Length - cursor - 12)) break;
                if (PrefixAt(bytes, cursor + 4, 4, new byte[] { 112, 72, 89, 115 })) {
                    if (length != 9) candidates.Add(new OfficeImageResolutionEvidence(true, "PNG/pHYs", null, false));
                    else Add("PNG/pHYs", OfficeExifProfileCodec.Read(bytes, cursor + 8, 4, false), OfficeExifProfileCodec.Read(bytes, cursor + 12, 4, false),
                        bytes[cursor + 16] == 0 ? OfficeImageResolutionUnit.AspectRatio : OfficeImageResolutionUnit.PixelsPerMeter, bytes[cursor + 16] <= 1);
                }
                cursor += (int)length + 12;
            }
        }
        if (bytes != null && format == OfficeImageFormat.Bmp && bytes.Length >= 46 && OfficeExifProfileCodec.Read(bytes, 14, 4, true) >= 40) {
            int x = unchecked((int)OfficeExifProfileCodec.Read(bytes, 38, 4, true)), y = unchecked((int)OfficeExifProfileCodec.Read(bytes, 42, 4, true));
            if (x != 0 || y != 0) Add("BMP/density", x, y, OfficeImageResolutionUnit.PixelsPerMeter, true);
        }
        if (bytes != null && format == OfficeImageFormat.Gif && bytes.Length >= 13 && bytes[12] != 0) Add("GIF/aspect", (bytes[12] + 15D) / 64D, 1D, OfficeImageResolutionUnit.AspectRatio, true);
        foreach (IGrouping<string, OfficeExifFieldEvidence> source in fields.GroupBy(field => field.Source)) {
            List<OfficeExifFieldEvidence> sourceFields = source.ToList();
            if (!sourceFields.Any(field => field.Directory == OfficeExifDirectory.Image && (field.TagId == 282 || field.TagId == 283 || field.TagId == 296))) continue;
            object? xv = UniqueValue(sourceFields, OfficeExifTag.XResolution), yv = UniqueValue(sourceFields, OfficeExifTag.YResolution), uv = UniqueValue(sourceFields, OfficeExifTag.ResolutionUnit);
            bool hasUnit = sourceFields.Any(field => field.Directory == OfficeExifDirectory.Image && field.TagId == 296);
            ushort unit = uv is ushort authoredUnit ? authoredUnit : (ushort)2;
            bool valid = xv is OfficeRational x && yv is OfficeRational y && x.Denominator != 0 && y.Denominator != 0 &&
                (!hasUnit || uv is ushort) && unit >= 1 && unit <= 3;
            Add(source.Key + "/density", xv is OfficeRational xr ? xr.ToDouble() : 0, yv is OfficeRational yr ? yr.ToDouble() : 0,
                unit == 1 ? OfficeImageResolutionUnit.AspectRatio : unit == 3 ? OfficeImageResolutionUnit.PixelsPerCentimeter : OfficeImageResolutionUnit.PixelsPerInch, valid);
        }
        bool unavailable = issues.Any(issue => issue == "exif:field-retention-limit" || issue == "exif:invalid-profile" || issue == "exif:profile-byte-limit" || issue == "exif:malformed-or-over-limit-carrier" || issue == "exif:bigtiff-profile-unsupported" || issue == "image:malformed-or-over-limit-profile");
        if (candidates.Count == 0) return new OfficeImageResolutionEvidence(false, unavailable ? "unavailable" : "none", null, false, !unavailable);
        if (unavailable) {
            OfficeXmpInspector.AddIssue(issues, "resolution:other-carrier-unavailable");
            return new OfficeImageResolutionEvidence(true, string.Join(";", candidates.Select(candidate => candidate.Source)), null, false);
        }
        if (candidates.Any(candidate => !candidate.IsValid)) {
            OfficeXmpInspector.AddIssue(issues, "resolution:invalid-or-incomplete-carrier");
            return new OfficeImageResolutionEvidence(true, string.Join(";", candidates.Select(candidate => candidate.Source)), null, false);
        }
        OfficeImageResolutionEvidence selected = candidates[0];
        if (candidates.Skip(1).Any(candidate => !SameResolution(selected.Resolution!, candidate.Resolution!))) {
            OfficeXmpInspector.AddIssue(issues, "resolution:conflicting-carriers");
            return new OfficeImageResolutionEvidence(true, string.Join(";", candidates.Select(candidate => candidate.Source)), null, false);
        }
        return selected;
        void Add(string source, double x, double y, OfficeImageResolutionUnit unit, bool valid) {
            bool acceptable = valid && x > 0 && y > 0 && !double.IsNaN(x) && !double.IsNaN(y) && !double.IsInfinity(x) && !double.IsInfinity(y);
            candidates.Add(new OfficeImageResolutionEvidence(true, source, acceptable ? new OfficeImageResolution(x, y, unit) : null, acceptable));
        }
    }
    private static bool SameResolution(OfficeImageResolution left, OfficeImageResolution right) {
        if (left.Unit == OfficeImageResolutionUnit.AspectRatio || right.Unit == OfficeImageResolutionUnit.AspectRatio) {
            return left.Unit == right.Unit && Close(left.Horizontal / left.Vertical, right.Horizontal / right.Vertical);
        }
        return Close(left.PhysicalDpiX!.Value, right.PhysicalDpiX!.Value) && Close(left.PhysicalDpiY!.Value, right.PhysicalDpiY!.Value);
    }
    private static bool Close(double left, double right) => Math.Abs(left - right) <= 1e-12D * Math.Max(1D, Math.Max(Math.Abs(left), Math.Abs(right)));
}

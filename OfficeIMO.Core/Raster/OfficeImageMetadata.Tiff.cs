using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeImageMetadata {
    private static void ReadTiffProfiles(byte[] input, OfficeImageMetadata metadata, CancellationToken token, long additionallyRetainedBytes = 0L) {
        long initialProfiles = metadata.RetainedProfileBytes;
        OfficeExifProfileCodec.Profile source = OfficeExifProfileCodec.Parse(input, imageTiff: true, cancellationToken: token, additionallyRetainedBytes: checked(additionallyRetainedBytes + initialProfiles));
        foreach (KeyValuePair<OfficeExifDirectory, OfficeExifProfileCodec.Directory> directory in source.Directories) {
            foreach (OfficeExifProfileCodec.Field entry in directory.Value.Fields) {
                token.ThrowIfCancellationRequested();
                if (IsStructural(entry.Tag.Id)) continue;
                if (checked(source.RetainedManagedBytes + Math.Max(0L, metadata.RetainedProfileBytes - initialProfiles) + entry.ValueLength * 4L + OfficeTiffPixelRanges.PlanningBytes) > OfficeRasterGuards.MaximumDecodedBytes) throw new FormatException("TIFF metadata copying exceeds the managed working-set limit.");
                if (directory.Key == OfficeExifDirectory.Image && IsTiffProfileTag(entry.Tag.Id)) {
                    byte[] profile = Slice(source.Bytes, entry.ValueOffset, entry.ValueLength);
                    if (entry.Tag.Id == 700) metadata.XmpProfile = profile;
                    if (entry.Tag.Id == 34675) metadata.IccProfile = profile;
                    if (entry.Tag.Id == 33723) metadata.IptcProfile = profile;
                    continue;
                }
                metadata.SetExifValue(entry.Tag, entry.Value, token);
                // MakerNotes can contain TIFF-relative private pointers. Retain their original entry
                // in same-container edits, and refuse implicit relocation in a profile export.
                if (directory.Key == OfficeExifDirectory.Exif && entry.Tag.Id == 37500) metadata._tiffOpaqueOffsets.Add(entry.Tag);
            }
        }
        OfficeExifValue? unit = metadata.GetExifValue(OfficeExifTag.ResolutionUnit);
        OfficeExifValue? horizontal = metadata.GetExifValue(OfficeExifTag.XResolution);
        OfficeExifValue? vertical = metadata.GetExifValue(OfficeExifTag.YResolution);
        if (horizontal?.Value is OfficeRational x && vertical?.Value is OfficeRational y && x.Denominator != 0 && y.Denominator != 0) {
            metadata.HorizontalResolution = x.ToDouble(); metadata.VerticalResolution = y.ToDouble();
            ushort value = unit?.Value is ushort number ? number : (ushort)2;
            metadata.ResolutionUnits = value == 3 ? OfficeImageResolutionUnit.PixelsPerCentimeter : value == 1 ? OfficeImageResolutionUnit.AspectRatio : OfficeImageResolutionUnit.PixelsPerInch;
        }
    }

    private static byte[] RewriteTiff(byte[] input, OfficeImageMetadata metadata, CancellationToken token, List<(long Start, long Length)>? pixelRangesOverride = null, bool preserveDensity = true, long additionallyRetainedBytes = 0L) {
        OfficeExifProfileCodec.Profile source = OfficeExifProfileCodec.Parse(input, imageTiff: true, cancellationToken: token, additionallyRetainedBytes: checked(additionallyRetainedBytes + metadata.RetainedProfileBytes));
        var changes = new Dictionary<OfficeExifTag, OfficeExifValue>();
        var removed = new HashSet<OfficeExifTag>();
        var target = new Dictionary<OfficeExifTag, OfficeExifValue>();
        foreach (OfficeExifValue value in metadata.ExifValues) {
            if (value.Tag.Directory == OfficeExifDirectory.Image && !OfficeExifProfileCodec.IsTiffMetadataTag(value.Tag.Id)) continue;
            target[value.Tag] = value;
        }
        if (preserveDensity) {
            GetExifResolution(metadata, out double x, out double y, out ushort unit);
            var xResolution = new OfficeRational(checked((uint)Math.Round(x * 1000D)), 1000);
            var yResolution = new OfficeRational(checked((uint)Math.Round(y * 1000D)), 1000);
            target[OfficeExifTag.XResolution] = new OfficeExifValue(OfficeExifTag.XResolution, xResolution);
            target[OfficeExifTag.YResolution] = new OfficeExifValue(OfficeExifTag.YResolution, yResolution);
            target[OfficeExifTag.ResolutionUnit] = new OfficeExifValue(OfficeExifTag.ResolutionUnit, unit);
        }
        AddProfile(700, OfficeExifDataType.Byte, metadata._xmp);
        AddProfile(34675, OfficeExifDataType.Undefined, metadata._icc);
        AddProfile(33723, OfficeExifDataType.Long, metadata._iptc);
        var original = new Dictionary<OfficeExifTag, OfficeExifProfileCodec.Field>();
        foreach (OfficeExifProfileCodec.Directory directory in source.Directories.Values) foreach (OfficeExifProfileCodec.Field entry in directory.Fields) {
            token.ThrowIfCancellationRequested();
            if (IsStructural(entry.Tag.Id)) continue;
            if (original.ContainsKey(entry.Tag)) throw new FormatException("Duplicate TIFF metadata tags cannot be edited safely.");
            original.Add(entry.Tag, entry);
            if (!target.ContainsKey(entry.Tag)) removed.Add(entry.Tag);
        }
        foreach (KeyValuePair<OfficeExifTag, OfficeExifValue> replacement in target) {
            if (original.TryGetValue(replacement.Key, out OfficeExifProfileCodec.Field? previous) && previous.Tag.DataType == replacement.Value.DataType && SameValue(previous.Value, replacement.Value.Value, token)) continue;
            if (metadata._tiffOpaqueOffsets.Contains(replacement.Key)) throw new NotSupportedException("TIFF-relative maker notes can only be preserved in their original TIFF container.");
            changes[replacement.Key] = replacement.Value;
        }
        if (!metadata.HasExifProfile) {
            foreach (KeyValuePair<OfficeExifDirectory, OfficeExifProfileCodec.Directory> directory in source.Directories) {
                if (directory.Key == OfficeExifDirectory.Image) continue;
                int count = (int)OfficeExifProfileCodec.Read(source.Bytes, directory.Value.Offset, 2, source.Little);
                for (int i = 0; i < count; i++) {
                    int at = directory.Value.Offset + 2 + i * 12;
                    int type = (int)OfficeExifProfileCodec.Read(source.Bytes, at + 2, 2, source.Little);
                    int id = (int)OfficeExifProfileCodec.Read(source.Bytes, at, 2, source.Little);
                    if (type == 13 || id == 330) throw new NotSupportedException("Removing TIFF Exif containing private IFD pointer arrays is not supported.");
                }
            }
            removed.Add(new OfficeExifTag(34665, OfficeExifDataType.Long));
            removed.Add(new OfficeExifTag(34853, OfficeExifDataType.Long));
        }
        long planningRetention = checked(source.RetainedManagedBytes - input.LongLength);
        List<(long Start, long Length)> pixelRanges = pixelRangesOverride ?? OfficeTiffPixelRanges.Read(input, token, planningRetention);
        foreach (KeyValuePair<OfficeExifTag, OfficeExifProfileCodec.Field> pair in original) {
            token.ThrowIfCancellationRequested();
            if (!changes.ContainsKey(pair.Key) && !removed.Contains(pair.Key) || pair.Value.ValueLength <= 4) continue;
            if (OfficeTiffPixelRanges.Overlaps(pixelRanges, pair.Value.ValueOffset, pair.Value.ValueLength)) throw new FormatException("TIFF metadata overlaps encoded pixel data and cannot be edited safely.");
        }
        if (changes.Count != 0 || removed.Count != 0) foreach (OfficeExifProfileCodec.Directory directory in source.Directories.Values) {
            token.ThrowIfCancellationRequested();
            int count = (int)OfficeExifProfileCodec.Read(source.Bytes, directory.Offset, 2, source.Little); long length = 6L + count * 12L;
            if (OfficeTiffPixelRanges.Overlaps(pixelRanges, directory.Offset, length)) throw new FormatException("A TIFF directory overlaps encoded pixel data and cannot be rewritten safely.");
        }
        byte[] result = OfficeExifProfileCodec.Encode(source, changes, removed, imageTiff: true, cancellationToken: token, additionallyRetainedBytes: OfficeTiffPixelRanges.RetainedBytes(pixelRanges));
        token.ThrowIfCancellationRequested(); return result;

        void AddProfile(ushort id, OfficeExifDataType type, byte[]? bytes) {
            if (bytes == null) return;
            object value = bytes;
            if (type == OfficeExifDataType.Long) {
                var words = new uint[(bytes.Length + 3) / 4];
                for (int i = 0; i < bytes.Length; i++) { if ((i & 4095) == 0) token.ThrowIfCancellationRequested(); words[i / 4] |= (uint)bytes[i] << (8 * (source.Little ? i % 4 : 3 - i % 4)); }
                value = words;
            }
            var tag = new OfficeExifTag(id, type); target[tag] = new OfficeExifValue(tag, value);
        }
    }

    private static bool IsTiffProfileTag(ushort id) => id == 700 || id == 34675 || id == 33723;
    private static bool SameValue(object first, object second, CancellationToken token) {
        if (first is Array left && second is Array right) {
            if (left.Length != right.Length || left.GetType() != right.GetType()) return false;
            for (int i = 0; i < left.Length; i++) { if ((i & 4095) == 0) token.ThrowIfCancellationRequested(); if (!Equals(left.GetValue(i), right.GetValue(i))) return false; }
            return true;
        }
        return Equals(first, second);
    }

}

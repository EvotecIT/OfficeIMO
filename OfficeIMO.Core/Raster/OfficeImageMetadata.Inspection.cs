using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Provenance;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeImageMetadata {
    /// <summary>Inspects supported encoded image metadata without editing the input or decoding pixels.</summary>
    /// <remarks>Classic Exif inspection covers the primary image, camera, GPS and interoperability directories.
    /// Additional TIFF pages, private pointer trees and MakerNote semantics are explicitly outside this profile.
    /// Malformed profiles produce diagnostics; unreadable image containers and exceeded input limits throw.</remarks>
    /// <exception cref="OfficeInspectionResourceLimitException">The configured metadata input-byte ceiling was exceeded.</exception>
    public static OfficeImageMetadataInspection Inspect(byte[] encodedBytes, OfficeImageMetadataInspectionOptions? options = null, CancellationToken cancellationToken = default) {
        if (encodedBytes == null) throw new ArgumentNullException(nameof(encodedBytes));
        OfficeImageMetadataInspectionOptions policy = (options ?? new OfficeImageMetadataInspectionOptions()).Snapshot();
        cancellationToken.ThrowIfCancellationRequested();
        if (encodedBytes.Length > policy.MaximumInputBytes) throw new OfficeInspectionResourceLimitException("image-metadata", nameof(policy.MaximumInputBytes), policy.MaximumInputBytes);
        if (!OfficeImageReader.TryIdentifyByContent(encodedBytes, null, cancellationToken, out OfficeImageInfo info)) throw new FormatException("The image metadata container is malformed or unsupported.");
        var fields = new List<OfficeExifFieldEvidence>(); var xmp = new List<OfficeXmpProperty>(); var issues = new List<string>();
        OfficeIccProfileInspection? icc = null;
        int exifCount = 0, iccCount = 0, xmpCount = 0;
        void Profile(string family, byte[]? bytes, int offset) {
            cancellationToken.ThrowIfCancellationRequested();
            string source = info.Format + "/" + family + "@" + offset;
            if (bytes == null) {
                if (family == "XMP") xmpCount++;
                if (family == "ICC") { iccCount++; icc = null; }
                OfficeXmpInspector.AddIssue(issues, "image:profile-byte-limit");
                return;
            }
            if (family == "Exif") { exifCount++; InspectExif(bytes, false, source, policy, fields, issues, cancellationToken); }
            else if (family == "XMP") { xmpCount++; OfficeXmpInspector.Read(bytes, source, policy, xmp, issues, cancellationToken); }
            else if (family == "ICC") {
                iccCount++;
                if (iccCount > 1) { OfficeXmpInspector.AddIssue(issues, "icc:duplicate-carriers"); icc = null; return; }
                icc = OfficeIccProfileInspector.Inspect(bytes, policy, cancellationToken);
                foreach (string issue in icc.Issues) OfficeXmpInspector.AddIssue(issues, issue);
            }
        }
        try {
            if (info.Format == OfficeImageFormat.Tiff) {
                if (encodedBytes.Length >= 4 && OfficeExifProfileCodec.Read(encodedBytes, 2, 2, encodedBytes[0] == 'I') == 43) issues.Add("exif:bigtiff-profile-unsupported");
                else {
                    InspectExif(encodedBytes, true, "TIFF/primary", policy, fields, issues, cancellationToken);
                    foreach (OfficeExifFieldEvidence field in fields) {
                        if (field.Directory != OfficeExifDirectory.Image || field.TagId != 700 && field.TagId != 34675) continue;
                        if (field.ValueLengthBytes > policy.MaximumProfileBytes) { OfficeXmpInspector.AddIssue(issues, "image:profile-byte-limit"); continue; }
                        Profile(field.TagId == 700 ? "XMP" : "ICC", Slice(encodedBytes, field.ValueOffset, field.ValueLengthBytes), field.EntryOffset);
                    }
                }
            } else if (info.Format == OfficeImageFormat.Jpeg) {
                for (int cursor = 2; cursor < encodedBytes.Length && OfficeProvenanceJpeg.TryReadMarker(encodedBytes, cursor, out byte marker, out int payload, out int length, out int end); cursor = end) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (marker == 0xDA || marker == 0xD9) break;
                    if (marker == 0xE1 && PrefixAt(encodedBytes, payload, length, JpegExifPrefix)) {
                        if (length > policy.MaximumProfileBytes) OfficeXmpInspector.AddIssue(issues, "exif:profile-byte-limit");
                        else Profile("Exif", Slice(encodedBytes, payload, length), cursor);
                    }
                }
                var snapshot = new OfficeImageMetadataSnapshot();
                OfficeImageMetadataInspector.InspectJpeg(encodedBytes, snapshot, 0L, cancellationToken, policy.MaximumProfileBytes);
                if (snapshot.Icc != null) Profile("ICC", snapshot.Icc, 0);
                else if ((snapshot.Kinds & OfficeImageMetadataKinds.Icc) != 0) OfficeXmpInspector.AddIssue(issues, "icc:unreadable-or-over-limit-carrier");
                var xmlPolicy = new OfficeProvenanceOptions { MaxAssetBytes = policy.MaximumProfileBytes, MaxContainerEntries = policy.MaximumXmlNodes, CancellationToken = cancellationToken };
                OfficeProvenanceJpegXmpResult packets = OfficeProvenanceJpegXmp.ProcessImage(encodedBytes, 2, 0, xmlPolicy, null, null, null,
                    (packet, source, valid) => {
                        xmpCount++;
                        if (!valid) OfficeXmpInspector.AddIssue(issues, "xmp:ambiguous-or-invalid-packet-set");
                        OfficeXmpInspector.Read(packet, source, policy, xmp, issues, cancellationToken);
                    });
                foreach (string issue in packets.InspectionIssues) OfficeXmpInspector.AddIssue(issues, issue);
            } else if (info.Format == OfficeImageFormat.Png) ReadPngProfiles(encodedBytes, null, cancellationToken, policy.MaximumProfileBytes, Profile, issue => OfficeXmpInspector.AddIssue(issues, issue));
            else if (info.Format == OfficeImageFormat.Webp) ReadWebpProfiles(encodedBytes, null, cancellationToken, policy.MaximumProfileBytes, Profile, issue => OfficeXmpInspector.AddIssue(issues, issue));
            else if (info.Format == OfficeImageFormat.Gif) {
                _ = OfficeProvenanceGif.RewriteMetadataProfiles(encodedBytes, null, null, OfficeImageMetadataProfileKinds.None,
                    false, cancellationToken, out _, out _, out _, readOnly: true,
                    profileObserver: (kind, bytes, offset) => Profile(kind == OfficeImageMetadataProfileKinds.Xmp ? "XMP" : "ICC", bytes, offset),
                    maximumRetainedProfileBytes: policy.MaximumProfileBytes);
            } else if (info.Format == OfficeImageFormat.Bmp) {
                var metadata = new OfficeImageMetadata();
                ReadBmpProfiles(encodedBytes, metadata, cancellationToken);
                if (metadata._xmp != null) Profile("XMP", metadata._xmp, 0);
                if (metadata._icc != null) Profile("ICC", metadata._icc, 0);
            } else issues.Add("image:profile-carriers-not-supported");
        } catch (Exception exception) when (exception is FormatException || exception is IOException || exception is NotSupportedException || exception is ArgumentException || exception is OverflowException) {
            OfficeXmpInspector.AddIssue(issues, "image:malformed-or-over-limit-profile");
        }
        if (exifCount > 1) OfficeXmpInspector.AddIssue(issues, "exif:duplicate-carriers");
        if (iccCount > 1) OfficeXmpInspector.AddIssue(issues, "icc:duplicate-carriers");
        if (xmpCount > 1 && info.Format != OfficeImageFormat.Jpeg) OfficeXmpInspector.AddIssue(issues, "xmp:duplicate-carriers");
        return CompleteInspection(info.Format, encodedBytes, fields, xmp, icc, issues, cancellationToken);
    }

    /// <summary>Inspects a standalone classic TIFF Exif profile, optionally including JPEG Exif framing.</summary>
    /// <exception cref="OfficeInspectionResourceLimitException">The configured metadata input-byte ceiling was exceeded.</exception>
    public static OfficeImageMetadataInspection InspectExifProfile(byte[] profile, OfficeImageMetadataInspectionOptions? options = null, CancellationToken cancellationToken = default) {
        if (profile == null) throw new ArgumentNullException(nameof(profile));
        OfficeImageMetadataInspectionOptions policy = (options ?? new OfficeImageMetadataInspectionOptions()).Snapshot();
        if (profile.Length > policy.MaximumInputBytes) throw new OfficeInspectionResourceLimitException("image-metadata", nameof(policy.MaximumInputBytes), policy.MaximumInputBytes);
        var fields = new List<OfficeExifFieldEvidence>(); var issues = new List<string>();
        InspectExif(profile, false, "Exif/profile", policy, fields, issues, cancellationToken);
        return CompleteInspection(OfficeImageFormat.Unknown, null, fields, new List<OfficeXmpProperty>(), null, issues, cancellationToken);
    }

    private static void InspectExif(byte[] bytes, bool imageTiff, string source, OfficeImageMetadataInspectionOptions policy,
        List<OfficeExifFieldEvidence> fields, List<string> issues, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (!imageTiff && bytes.Length > policy.MaximumProfileBytes) { OfficeXmpInspector.AddIssue(issues, "exif:profile-byte-limit"); return; }
        try {
            OfficeExifProfileCodec.Profile profile = OfficeExifProfileCodec.Parse(bytes, imageTiff, token, inspectionFields: fields, inspectionOptions: policy, inspectionSource: source);
            if (profile.OmittedInspectionFields != 0) OfficeXmpInspector.AddIssue(issues, "exif:field-retention-limit");
            foreach (OfficeExifProfileCodec.Directory directory in profile.Directories.Values) if (directory.Next != 0) OfficeXmpInspector.AddIssue(issues, "exif:additional-ifd-values-not-inspected");
        } catch (Exception exception) when (exception is FormatException || exception is OverflowException) { OfficeXmpInspector.AddIssue(issues, "exif:invalid-profile"); }
    }

    private static OfficeImageMetadataInspection CompleteInspection(OfficeImageFormat format, byte[]? encodedBytes,
        List<OfficeExifFieldEvidence> fields, List<OfficeXmpProperty> xmp, OfficeIccProfileInspection? icc, List<string> issues, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        foreach (OfficeExifFieldEvidence field in fields) {
            if (field.Omission != null && field.Omission != "exif:structural-pointer") OfficeXmpInspector.AddIssue(issues, field.Omission);
            if (field.Directory == OfficeExifDirectory.Exif && field.TagId == 37500) OfficeXmpInspector.AddIssue(issues, "exif:makernote-semantics-not-inspected");
            if (field.RawDataType == 13 || field.TagId == 330) OfficeXmpInspector.AddIssue(issues, "exif:private-ifd-values-not-inspected");
        }
        if (fields.GroupBy(field => Tuple.Create(field.Source, field.Directory, field.TagId)).Any(group => group.Count() > 1)) OfficeXmpInspector.AddIssue(issues, "exif:duplicate-tag");
        var timestamps = new List<OfficeExifTimestamp>();
        ReadTimestamp(OfficeExifTag.DateTime, OfficeExifTag.OffsetTime, OfficeExifTag.SubSecTime);
        ReadTimestamp(OfficeExifTag.DateTimeOriginal, OfficeExifTag.OffsetTimeOriginal, OfficeExifTag.SubSecTimeOriginal);
        ReadTimestamp(OfficeExifTag.DateTimeDigitized, OfficeExifTag.OffsetTimeDigitized, OfficeExifTag.SubSecTimeDigitized);
        string[] gpsSources = fields.Where(field => field.Directory == OfficeExifDirectory.Gps).Select(field => field.Source).Distinct().ToArray();
        List<OfficeExifFieldEvidence> gpsFields = gpsSources.Length == 1 ? fields.Where(field => field.Source == gpsSources[0]).ToList() : new List<OfficeExifFieldEvidence>();
        if (HasDuplicatePointer(gpsFields, OfficeExifDirectory.Image, 34853)) {
            OfficeXmpInspector.AddIssue(issues, "exif:directory-pointer-ambiguous"); gpsFields.Clear();
        }
        // Unretained entries may include another pointer or coordinate component.
        // A complete-looking retained tuple cannot establish source uniqueness.
        if (issues.Contains("exif:field-retention-limit")) gpsFields.Clear();
        object? latitude = UniqueValue(gpsFields, OfficeExifTag.GPSLatitude), longitude = UniqueValue(gpsFields, OfficeExifTag.GPSLongitude);
        string? latitudeRef = UniqueValue(gpsFields, OfficeExifTag.GPSLatitudeRef) as string, longitudeRef = UniqueValue(gpsFields, OfficeExifTag.GPSLongitudeRef) as string;
        OfficeExifGpsCoordinate.TryCreate(latitude as OfficeRational[], latitudeRef, longitude as OfficeRational[], longitudeRef, out OfficeExifGpsCoordinate? gps);
        if (gps == null && fields.Any(field => field.Directory == OfficeExifDirectory.Gps && field.TagId >= 1 && field.TagId <= 4)) OfficeXmpInspector.AddIssue(issues, "exif:gps-incomplete-invalid-or-ambiguous");
        OfficeImageResolutionEvidence resolution = InspectResolution(format, encodedBytes, fields, issues, token);
        return new OfficeImageMetadataInspection(format, fields, timestamps, gps, resolution, xmp, icc, issues);
        void ReadTimestamp(OfficeExifTag date, OfficeExifTag offset, OfficeExifTag fraction) {
            foreach (IGrouping<string, OfficeExifFieldEvidence> source in fields.GroupBy(field => field.Source)) {
                List<OfficeExifFieldEvidence> localFields = source.ToList();
                OfficeExifFieldEvidence[] matches = localFields.Where(field => field.Directory == date.Directory && field.TagId == date.Id).ToArray();
                if (matches.Length == 0) continue;
                if (HasDuplicatePointer(localFields, OfficeExifDirectory.Image, 34665) || HasDuplicatePointer(localFields, OfficeExifDirectory.Image, 34853) || HasDuplicatePointer(localFields, OfficeExifDirectory.Exif, 40965)) {
                    OfficeXmpInspector.AddIssue(issues, "exif:directory-pointer-ambiguous"); continue;
                }
                if (issues.Contains("exif:field-retention-limit")) { OfficeXmpInspector.AddIssue(issues, "exif:timestamp-invalid-or-ambiguous"); continue; }
                if (matches.Length != 1 || !(matches[0].Value?.Value is string raw)) { OfficeXmpInspector.AddIssue(issues, "exif:timestamp-invalid-or-ambiguous"); continue; }
                if (HasUnusable(localFields, offset) || HasUnusable(localFields, fraction)) { OfficeXmpInspector.AddIssue(issues, "exif:timestamp-companion-invalid-or-ambiguous"); continue; }
                OfficeExifTimestamp value = OfficeExifTimestamp.Parse(raw, UniqueValue(localFields, offset) as string, UniqueValue(localFields, fraction) as string);
                value.FieldName = date.Name; value.Source = matches[0].Source; timestamps.Add(value);
                if (value.Issue != null) OfficeXmpInspector.AddIssue(issues, value.Issue);
            }
        }
        bool HasUnusable(List<OfficeExifFieldEvidence> sourceFields, OfficeExifTag tag) => sourceFields.Any(field => field.Directory == tag.Directory && field.TagId == tag.Id) && !(UniqueValue(sourceFields, tag) is string);
    }

    private static bool HasDuplicatePointer(List<OfficeExifFieldEvidence> fields, OfficeExifDirectory directory, ushort tagId) =>
        fields.Count(field => field.Directory == directory && field.TagId == tagId) > 1;

    private static object? UniqueValue(List<OfficeExifFieldEvidence> fields, OfficeExifTag tag) {
        OfficeExifFieldEvidence? found = null;
        foreach (OfficeExifFieldEvidence field in fields) {
            if (field.Directory != tag.Directory || field.TagId != tag.Id) continue;
            if (found != null) return null;
            found = field;
        }
        return found?.Value?.Value;
    }
}

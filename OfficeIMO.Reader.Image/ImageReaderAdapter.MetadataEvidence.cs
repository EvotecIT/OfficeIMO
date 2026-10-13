using System.Collections;
using OfficeIMO.Drawing;

namespace OfficeIMO.Reader.Image;

internal static partial class ImageReaderAdapter {
    private static void ProjectMetadataEvidence(byte[] bytes, OfficeDocumentReadResult result, ReaderOptions readerOptions,
        ReaderImageOptions imageOptions, ReaderLocation location, CancellationToken token) {
        var builder = new ReaderInspectionMetadataBuilder(readerOptions, "image", token);
        OfficeImageMetadataInspection evidence;
        try { evidence = OfficeImageMetadata.Inspect(bytes, imageOptions.MetadataInspection, token); }
        catch (Exception exception) when (exception is FormatException || exception is InvalidDataException) {
            builder.Coverage("ImageMetadata", "unavailable", "The encoded image or owner metadata budget does not permit this inspection profile.", location);
            builder.Diagnostic("IMAGE_METADATA_INSPECTION_UNAVAILABLE", "Detailed image metadata could not be inspected within the supported container and byte limits.", location);
            builder.Complete(result);
            return;
        }
        builder.Coverage("ImageMetadata", evidence.IsComplete ? "complete" : "partial",
            "Primary classic-TIFF Exif directories, namespace-qualified XMP values and bounded ICC headers/tags. Pixels, private MakerNote semantics and additional TIFF IFD values are outside this profile.", location);
        builder.Add("image.resolution", "AuthoredResolution", evidence.Resolution.IsPresenceKnown ? evidence.Resolution.IsDeclared ? "true" : "false" : "unknown", "inspection-presence", location,
            attributes: ReaderInspectionMetadataBuilder.Attributes(("source", evidence.Resolution.Source), ("valid", evidence.Resolution.IsValid ? "true" : "false")));
        if (evidence.Resolution.Resolution is OfficeImageResolution resolution) {
            builder.Add("image.resolution", "NativeResolution", resolution.Horizontal.ToString("R", CultureInfo.InvariantCulture) + "," + resolution.Vertical.ToString("R", CultureInfo.InvariantCulture), "number-pair", location,
                attributes: ReaderInspectionMetadataBuilder.Attributes(("unit", resolution.Unit.ToString()), ("source", evidence.Resolution.Source)));
        }
        foreach (OfficeExifFieldEvidence field in evidence.ExifFields) {
            token.ThrowIfCancellationRequested();
            builder.Add("image.exif", "0x" + field.TagId.ToString("X4", CultureInfo.InvariantCulture), FormatValue(field.Value?.Value), "exif-value", location,
                attributes: ReaderInspectionMetadataBuilder.Attributes(("directory", field.Directory.ToString()), ("source", field.Source),
                    ("type", field.RawDataType.ToString(CultureInfo.InvariantCulture)), ("count", field.Count.ToString(CultureInfo.InvariantCulture)),
                    ("entryOffset", field.EntryOffset.ToString(CultureInfo.InvariantCulture)), ("valueOffset", field.ValueOffset.ToString(CultureInfo.InvariantCulture)),
                    ("offsetBasis", "TIFF-profile"), ("valueLengthBytes", field.ValueLengthBytes.ToString(CultureInfo.InvariantCulture)), ("omission", field.Omission)));
        }
        foreach (OfficeExifTimestamp timestamp in evidence.Timestamps) {
            builder.Add("image.exif.timestamp", timestamp.FieldName, timestamp.Instant?.ToString("O", CultureInfo.InvariantCulture) ?? timestamp.LocalDateTime?.ToString("O", CultureInfo.InvariantCulture),
                timestamp.HasKnownTimezone ? "date-time-offset" : "local-date-time", location,
                attributes: ReaderInspectionMetadataBuilder.Attributes(("raw", timestamp.RawValue), ("rawOffset", timestamp.RawOffset), ("rawSubseconds", timestamp.RawSubseconds),
                    ("timezone", timestamp.HasKnownTimezone ? "explicit-offset" : "unknown"), ("valid", timestamp.IsValid ? "true" : "false"), ("source", timestamp.Source), ("issue", timestamp.Issue)));
        }
        if (evidence.GpsCoordinate is OfficeExifGpsCoordinate gps) builder.Add("image.exif.gps", "Coordinates",
            gps.Latitude.ToString("R", CultureInfo.InvariantCulture) + "," + gps.Longitude.ToString("R", CultureInfo.InvariantCulture), "decimal-degrees", location);
        foreach (OfficeXmpProperty property in evidence.XmpProperties) {
            builder.Add("image.xmp", property.Name, property.Value, "xmp-value", location,
                attributes: ReaderInspectionMetadataBuilder.Attributes(("namespace", property.NamespaceUri), ("kind", property.Kind.ToString()), ("resource", property.IsResource ? "true" : "false"),
                    ("language", property.Language), ("index", property.ItemIndex?.ToString(CultureInfo.InvariantCulture)), ("source", property.Source),
                    ("truncated", property.IsTruncated ? "true" : "false")));
        }
        if (evidence.IccProfile is OfficeIccProfileInspection icc) {
            builder.Add("image.icc", "Profile", icc.ProfileClassSignature, "icc-profile", location,
                attributes: ReaderInspectionMetadataBuilder.Attributes(("structurallyValid", icc.IsStructurallyValid ? "true" : "false"),
                    ("supportedInputTransform", icc.HasSupportedInputTransform?.ToString().ToLowerInvariant()), ("deviceColorSpace", icc.DeviceColorSpaceSignature),
                    ("connectionSpace", icc.ConnectionSpaceSignature), ("declaredLengthBytes", icc.DeclaredLengthBytes.ToString(CultureInfo.InvariantCulture)),
                    ("declaredTagCount", icc.DeclaredTagCount.ToString(CultureInfo.InvariantCulture))));
            foreach (OfficeIccTagEvidence tag in icc.Tags) builder.Add("image.icc.tag", tag.Signature, tag.TypeSignature, "icc-tag", location,
                attributes: ReaderInspectionMetadataBuilder.Attributes(("offset", tag.Offset.ToString(CultureInfo.InvariantCulture)), ("lengthBytes", tag.LengthBytes.ToString(CultureInfo.InvariantCulture)), ("offsetBasis", "ICC-profile")));
        }
        foreach (string issue in evidence.Issues) builder.Diagnostic("IMAGE_METADATA_" + issue.Replace(':', '_').Replace('-', '_').ToUpperInvariant(), issue, location);
        builder.Complete(result);
    }
    private static string? FormatValue(object? value) {
        if (value == null) return null;
        if (value is byte[] bytes) return BitConverter.ToString(bytes).Replace("-", string.Empty);
        if (value is IEnumerable sequence && !(value is string)) return string.Join(",", sequence.Cast<object>().Select(item => Convert.ToString(item, CultureInfo.InvariantCulture)));
        return Convert.ToString(value, CultureInfo.InvariantCulture);
    }
}

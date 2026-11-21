using OfficeIMO.Drawing;
using OfficeIMO.Provenance;

namespace OfficeIMO.Core.Internal;

/// <summary>
/// Hidden data carried by package assets, shared by every format that implements hidden-data removal: provenance (Content Credentials and
/// IPTC digital source types) and Exif, GPS, XMP and IPTC metadata in images. Provenance carriers are found and removed by the provenance
/// engine, so hidden-data removal and <c>RemoveProvenance</c> agree. Formats own only the package plumbing: reading and replacing parts.
/// </summary>
internal static class OfficeHiddenDataAssets {
    /// <summary>The ZIP entry name of package-level Content Credentials in Open XML, OpenDocument and EPUB packages.</summary>
    internal const string PackageManifestPath = "META-INF/content_credential.c2pa";

    private const string UnusedSuffix = " (picture not used by the document)";

    /// <summary>
    /// Reports and, when removing, cleans one image. Returns the replacement bytes, or null when the image is unchanged.
    /// Metadata is stripped in place without re-encoding pixels. <paramref name="part"/> is the package part or entry that holds the image;
    /// <paramref name="unused"/> marks a picture stored in the file that nothing in the document references.
    /// </summary>
    internal static byte[]? ProcessImage(byte[] original, string part, OfficeHiddenDataCollector collector, bool unused = false) {
        List<OfficeProvenanceEvidence> provenance = InspectImageProvenance(original, part);
        string? description = DescribePrivateImageMetadata(original);
        if (provenance.Count == 0 && description == null) return null;
        string suffix = unused ? UnusedSuffix : string.Empty;

        bool removeProvenance = collector.Removes(OfficeHiddenDataKind.Provenance);
        OfficeProvenanceRemovalOptions? carriers = removeProvenance ? collector.Removal!.Provenance : null;
        bool keepSourceType = carriers?.RemoveAiSourceMetadata != true;
        byte[] current = original;
        if (carriers != null && provenance.Count > 0) current = TryRemoveImageProvenance(current, part, carriers) ?? current;

        bool metadataRemoved = false;
        OfficeHiddenDataReason metadataReason = OfficeHiddenDataReason.None;
        if (description != null && collector.Removes(OfficeHiddenDataKind.ImageMetadata)) {
            List<OfficeProvenanceEvidence> currentProvenance = ReferenceEquals(current, original) ? provenance : InspectImageProvenance(current, part);
            bool bound = currentProvenance.Any(IsContentBinding);
            byte[]? stripped = bound ? null : TryStripPrivateImageMetadata(current, keepSourceType);
            if (stripped == null) {
                metadataReason = bound ? OfficeHiddenDataReason.ProtectsContentCredentials : OfficeHiddenDataReason.UnsupportedImageFormat;
            } else {
                current = stripped;
                metadataRemoved = true;
                if (keepSourceType && provenance.Any(item => item.Carrier == OfficeProvenanceCarrierKind.IptcDigitalSourceType)) {
                    metadataReason = OfficeHiddenDataReason.SourceTypeKept;
                }
            }
        }

        if (description != null) {
            OfficeHiddenDataOutcome outcome = collector.Resolve(OfficeHiddenDataKind.ImageMetadata, metadataRemoved, metadataReason, out OfficeHiddenDataReason reason);
            collector.Record(OfficeHiddenDataKind.ImageMetadata, part, description + suffix, null, outcome, null,
                scope: OfficeHiddenDataScope.EmbeddedImage, reason: metadataRemoved ? metadataReason : reason, part: part);
        }

        List<OfficeProvenanceEvidence> remaining = ReferenceEquals(current, original) ? provenance.ToList() : InspectImageProvenance(current, part);
        foreach (OfficeProvenanceEvidence evidence in provenance) {
            bool removed = !TakeMatch(remaining, evidence);
            OfficeHiddenDataOutcome outcome = collector.Resolve(OfficeHiddenDataKind.Provenance, removed, OfficeHiddenDataReason.KeptByProvenanceChoices,
                out OfficeHiddenDataReason reason);
            if (!removed && collector.Removal != null && IsContentBinding(evidence) && !ReferenceEquals(current, original)) {
                reason = OfficeHiddenDataReason.ContentCredentialsNoLongerVerify;
            }
            collector.Record(OfficeHiddenDataKind.Provenance, part + "/" + evidence.Location, DescribeProvenance(evidence) + suffix,
                ProvenanceValue(evidence), outcome, null, evidence, OfficeHiddenDataScope.EmbeddedImage, reason, part);
        }
        return ReferenceEquals(current, original) ? null : current;
    }

    /// <summary>
    /// Reports package-level Content Credentials. They bind every byte of the package, so no edit of a loaded document can keep them valid;
    /// removal always deletes them and says so. Returns whether the caller must delete the entry.
    /// </summary>
    internal static bool ProcessPackageManifest(byte[] manifest, string part, OfficeHiddenDataCollector collector) {
        var limits = new OfficeProvenanceOptions();
        bool valid = manifest.LongLength <= limits.MaxManifestBytes &&
            OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, limits.MaxManifestBytes, limits.MaxContainerEntries, out _);
        OfficeProvenanceEvidence evidence = new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.C2paManifest,
            "ZIP/" + PackageManifestPath + "[0]", valid, manifest.Length)
            .WithManifest(valid ? OfficeC2paManifestStore.TryDescribe(manifest, 0, manifest.Length) : null);
        bool remove = collector.Removal != null;
        OfficeHiddenDataReason reason = remove && !collector.Removes(OfficeHiddenDataKind.Provenance)
            ? OfficeHiddenDataReason.WholeFileCredentialsCannotSurvive
            : OfficeHiddenDataReason.None;
        collector.Record(OfficeHiddenDataKind.Provenance, part, DescribeProvenance(evidence), ProvenanceValue(evidence),
            remove ? OfficeHiddenDataOutcome.Removed : OfficeHiddenDataOutcome.Found, null, evidence, OfficeHiddenDataScope.Package, reason, part);
        return remove;
    }

    internal static byte[] ReadAll(Stream stream) {
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static List<OfficeProvenanceEvidence> InspectImageProvenance(byte[] data, string location) {
        var options = new OfficeProvenanceOptions { ProcessEmbeddedAssets = false };
        try {
            if (data.LongLength > options.MaxAssetBytes || !IsProvenanceImage(OfficeProvenanceInspector.DetectFormat(data, null, options))) {
                return new List<OfficeProvenanceEvidence>();
            }
            return OfficeProvenanceInspector.Inspect(data, location, options).Evidence.ToList();
        } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException or NotSupportedException or FormatException) {
            return new List<OfficeProvenanceEvidence>();
        }
    }

    private static byte[]? TryRemoveImageProvenance(byte[] data, string location, OfficeProvenanceRemovalOptions carriers) {
        var options = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = carriers.RemoveC2paManifests,
            RemoveExternalC2paReferences = carriers.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = carriers.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = carriers.RequireStructurallyValidCarrier,
            ProcessEmbeddedAssets = false
        };
        options.Limits.MaxAssetBytes = carriers.Limits.MaxAssetBytes;
        options.Limits.MaxManifestBytes = carriers.Limits.MaxManifestBytes;
        options.Limits.MaxCarriers = carriers.Limits.MaxCarriers;
        options.Limits.MaxContainerEntries = carriers.Limits.MaxContainerEntries;
        options.Limits.MaxExpandedContainerBytes = carriers.Limits.MaxExpandedContainerBytes;
        try {
            OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(data, location, options);
            return result.WasChanged ? result.ToArray() : null;
        } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException or NotSupportedException or FormatException) {
            return null;
        }
    }

    private static bool IsProvenanceImage(OfficeProvenanceAssetFormat format) =>
        format is OfficeProvenanceAssetFormat.Jpeg or OfficeProvenanceAssetFormat.Png or OfficeProvenanceAssetFormat.Webp or
            OfficeProvenanceAssetFormat.Gif or OfficeProvenanceAssetFormat.Tiff or OfficeProvenanceAssetFormat.Svg;

    /// <summary>Carriers whose signature covers the image bytes, so any metadata edit breaks them.</summary>
    private static bool IsContentBinding(OfficeProvenanceEvidence evidence) =>
        evidence.Carrier is OfficeProvenanceCarrierKind.C2paManifest or OfficeProvenanceCarrierKind.C2paExternalManifest;

    private static bool TakeMatch(List<OfficeProvenanceEvidence> remaining, OfficeProvenanceEvidence evidence) {
        int index = remaining.FindIndex(item => item.Carrier == evidence.Carrier && item.Value == evidence.Value &&
            item.PayloadLength == evidence.PayloadLength);
        if (index < 0) return false;
        remaining.RemoveAt(index);
        return true;
    }

    internal static string DescribeProvenance(OfficeProvenanceEvidence evidence) {
        string description = evidence.Carrier switch {
            OfficeProvenanceCarrierKind.C2paManifest => evidence.Manifest?.DeclaresGenerativeAi == true
                ? "Content Credentials (C2PA manifest) declaring generative AI"
                : "Content Credentials (C2PA manifest)",
            OfficeProvenanceCarrierKind.C2paExternalManifest => "Reference to external Content Credentials",
            _ => "Digital source type: " + evidence.DigitalSourceKind
        };
        return evidence.IsStructurallyValid ? description : description + " (malformed)";
    }

    internal static string? ProvenanceValue(OfficeProvenanceEvidence evidence) =>
        evidence.Carrier == OfficeProvenanceCarrierKind.C2paManifest ? evidence.Manifest?.ClaimGenerator : evidence.Value;

    private static string? DescribePrivateImageMetadata(byte[] data) {
        OfficeImageMetadata metadata;
        try {
            metadata = OfficeImageMetadata.Read(data);
        } catch (Exception exception) when (exception is FormatException or NotSupportedException or ArgumentException) {
            return null;
        }

        var profiles = new List<string>();
        bool hasGps = false;
        bool hasExif = false;
        foreach (OfficeExifValue value in metadata.ExifValues) {
            if (value.Tag.Directory == OfficeExifDirectory.Gps) hasGps = true;
            else if (!value.Tag.Equals(OfficeExifTag.Orientation)) hasExif = true;
        }

        if (hasExif) profiles.Add("Exif");
        if (hasGps) profiles.Add("GPS location");
        byte[]? xmp = metadata.XmpProfileData;
        if (xmp != null && !OfficeProvenanceXmp.HasOnlyDigitalSourceTypes(xmp, new OfficeProvenanceOptions())) profiles.Add("XMP");
        if (metadata.IptcProfile != null) profiles.Add("IPTC");
        return profiles.Count == 0 ? null : "Image metadata: " + string.Join(", ", profiles);
    }

    /// <summary>
    /// Rewrites only the metadata containers; ICC colour, density and orientation stay so the image looks the same.
    /// With <paramref name="keepDigitalSourceType"/>, XMP is reduced to the image's digital source type declaration.
    /// </summary>
    private static byte[]? TryStripPrivateImageMetadata(byte[] data, bool keepDigitalSourceType) {
        try {
            OfficeImageMetadata metadata = OfficeImageMetadata.Read(data);
            OfficeExifValue? orientation = metadata.GetExifValue(OfficeExifTag.Orientation);
            OfficeImageMetadata stripped = metadata.Clone();
            stripped.ClearExif();
            byte[]? xmp = metadata.XmpProfileData;
            stripped.XmpProfile = keepDigitalSourceType && xmp != null
                ? OfficeProvenanceXmp.CreateDigitalSourceTypePacket(xmp, new OfficeProvenanceOptions())
                : null;
            stripped.IptcProfile = null;
            if (orientation != null) stripped.SetExifValue(OfficeExifTag.Orientation, orientation.Value);
            return OfficeImageMetadata.Apply(data, stripped);
        } catch (Exception exception) when (exception is FormatException or NotSupportedException or ArgumentException) {
            return null;
        }
    }
}

using System;
using OfficeIMO.Provenance;

namespace OfficeIMO.ContentSafety;

/// <summary>Maps content-safety resource policy into the provenance pass used for signature handling.</summary>
internal static class OfficeContentSafetyProvenanceOptions {
    internal static OfficeProvenanceRemovalOptions CreateSignatureRemovalOptions(
        OfficeContentCleanupOptions cleanupOptions) {
        if (cleanupOptions == null) throw new ArgumentNullException(nameof(cleanupOptions));
        var provenanceOptions = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = cleanupOptions.SignatureMutationPolicy,
            MaxOutputBytes = cleanupOptions.Inspection.MaxInputBytes
        };
        provenanceOptions.Limits.MaxAssetBytes = cleanupOptions.Inspection.MaxInputBytes;
        provenanceOptions.Limits.MaxManifestBytes = Math.Min(
            provenanceOptions.Limits.MaxManifestBytes,
            cleanupOptions.Inspection.MaxInputBytes);
        provenanceOptions.Limits.MaxContainerEntries = cleanupOptions.Inspection.MaxPackageEntries;
        provenanceOptions.Limits.MaxExpandedContainerBytes = cleanupOptions.Inspection.MaxExpandedPackageBytes;
        return provenanceOptions;
    }
}

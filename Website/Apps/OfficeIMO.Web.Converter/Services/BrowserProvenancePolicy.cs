using OfficeIMO.Provenance;

namespace OfficeIMO.Web.Converter.Services;

internal static class BrowserProvenancePolicy {
    internal static OfficeProvenanceOptions Limits() => new() {
        MaxAssetBytes = BrowserConversionService.MaxPackageBytes, MaxManifestBytes = 4 * 1024 * 1024,
        MaxCarriers = 128, MaxContainerEntries = 16_384, MaxExpandedContainerBytes = 64 * 1024 * 1024,
        MaxEmbeddedAssets = 256, ProcessEmbeddedAssets = true
    };
    internal static OfficeProvenanceRemovalOptions Removal(bool manifests, bool references, bool declarations) => new() {
        RemoveC2paManifests = manifests, RemoveExternalC2paReferences = references, RemoveAiSourceMetadata = declarations,
        RequireStructurallyValidCarrier = true, SignatureMutationPolicy = OfficeIMO.OfficeSignatureMutationPolicy.BlockSave,
        Limits = { MaxAssetBytes = BrowserConversionService.MaxPackageBytes, MaxManifestBytes = 4 * 1024 * 1024,
            MaxCarriers = 128, MaxContainerEntries = 16_384, MaxExpandedContainerBytes = 64 * 1024 * 1024,
            MaxEmbeddedAssets = 256, ProcessEmbeddedAssets = true },
        MaxOutputBytes = BrowserConversionService.MaxPackageBytes, MaxEmbeddedAssets = 256, ProcessEmbeddedAssets = true
    };
}

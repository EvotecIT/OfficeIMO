using System;
using System.IO;
using System.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Provenance;

internal readonly struct OfficeProvenanceSignatureStripResult {
    internal OfficeProvenanceSignatureStripResult(byte[] data, bool hadSignatures) {
        Data = data ?? throw new ArgumentNullException(nameof(data));
        HadSignatures = hadSignatures;
    }
    internal byte[] Data { get; }
    internal bool HadSignatures { get; }
}

internal static class OfficeProvenancePackageMutation {
    internal static OfficeProvenanceRemovalResult RemoveFile(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options,
        Func<byte[], OfficeProvenanceOptions, OfficeProvenanceSignatureStripResult> stripSignatures,
        Func<byte[], bool>? hasSignatures = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeProvenanceRemovalOptions();
        string fullInputPath = Path.GetFullPath(inputPath);
        byte[] data;
        using (var stream = File.OpenRead(fullInputPath)) data = OfficeProvenanceBinary.ReadBounded(stream, options.Limits.MaxAssetBytes);
        OfficeProvenanceRemovalResult result = Remove(data, fullInputPath, options, stripSignatures, hasSignatures);
        OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), result.ToArray());
        return result;
    }

    internal static OfficeProvenanceRemovalResult Remove(
        byte[] data,
        string fileName,
        OfficeProvenanceRemovalOptions? options,
        Func<byte[], OfficeProvenanceOptions, OfficeProvenanceSignatureStripResult> stripSignatures,
        Func<byte[], bool>? hasSignatures = null) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (stripSignatures == null) throw new ArgumentNullException(nameof(stripSignatures));
        options ??= new OfficeProvenanceRemovalOptions();
        if (options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.PreserveSignatureMarkup) {
            return OfficeProvenanceRemover.Remove(data, fileName, options);
        }

        OfficeProvenanceRemovalOptions previewOptions = Clone(options, OfficeSignatureMutationPolicy.PreserveSignatureMarkup);
        OfficeProvenanceRemovalResult preview = OfficeProvenanceRemover.Remove(data, fileName, previewOptions);
        if (!preview.WasChanged) return preview;

        OfficeProvenanceZip.ValidateForOwningPackageMutation(data, options.Limits);
        bool hadSignatureEvidence = hasSignatures?.Invoke(data) ?? OfficeProvenanceZip.HasPackageSignature(data, options);
        if (options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.BlockSave) {
            if (hadSignatureEvidence) {
                throw new InvalidOperationException("Removing provenance would invalidate package signatures. Choose an explicit signature mutation policy.");
            }
            return preview;
        }
        OfficeProvenanceSignatureStripResult stripped = stripSignatures(data, options.Limits);
        if (!hadSignatureEvidence && !stripped.HadSignatures) return preview;
        if (hadSignatureEvidence && !stripped.HadSignatures) {
            throw new InvalidOperationException("The package contains signature evidence that its owning adapter could not remove safely.");
        }
        if (stripped.Data.SequenceEqual(data)) {
            throw new InvalidOperationException("The document reports signatures, but its owning package adapter could not remove them safely.");
        }
        if (hasSignatures?.Invoke(stripped.Data) ?? OfficeProvenanceZip.HasPackageSignature(stripped.Data, options)) {
            throw new InvalidOperationException("The owning package adapter left signature evidence in the rewritten document.");
        }
        OfficeProvenanceRemovalResult final = OfficeProvenanceRemover.Remove(stripped.Data, fileName, previewOptions);
        return new OfficeProvenanceRemovalResult(
            final.ToArray(),
            preview.Before,
            final.After,
            final.Changes,
            final.WasReserialized,
            wereInvalidatedSignaturesRemoved: true);
    }

    private static OfficeProvenanceRemovalOptions Clone(
        OfficeProvenanceRemovalOptions source,
        OfficeSignatureMutationPolicy signaturePolicy) {
        var clone = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = source.RemoveC2paManifests,
            RemoveExternalC2paReferences = source.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = source.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = source.RequireStructurallyValidCarrier,
            SignatureMutationPolicy = signaturePolicy,
            ProcessEmbeddedAssets = source.ProcessEmbeddedAssets && source.Limits.ProcessEmbeddedAssets,
            MaxEmbeddedAssets = Math.Min(source.MaxEmbeddedAssets, source.Limits.MaxEmbeddedAssets)
        };
        clone.Limits.MaxAssetBytes = source.Limits.MaxAssetBytes;
        clone.Limits.MaxManifestBytes = source.Limits.MaxManifestBytes;
        clone.Limits.MaxCarriers = source.Limits.MaxCarriers;
        clone.Limits.MaxContainerEntries = source.Limits.MaxContainerEntries;
        clone.Limits.MaxExpandedContainerBytes = source.Limits.MaxExpandedContainerBytes;
        clone.Limits.ProcessEmbeddedAssets = source.ProcessEmbeddedAssets && source.Limits.ProcessEmbeddedAssets;
        clone.Limits.MaxEmbeddedAssets = Math.Min(source.MaxEmbeddedAssets, source.Limits.MaxEmbeddedAssets);
        return clone;
    }
}

using System;
using System.IO;
using System.IO.Compression;
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
    /// <summary>Reads a bounded package, validates ownership, and inspects the same bytes for provenance.</summary>
    internal static OfficeProvenanceReport InspectFile(
        string filePath,
        OfficeProvenanceOptions? options,
        Action<byte[], OfficeProvenanceOptions> validatePackage) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        if (validatePackage == null) throw new ArgumentNullException(nameof(validatePackage));
        options ??= new OfficeProvenanceOptions();
        OfficeProvenanceBinary.ValidateLimits(options);
        string fullPath = Path.GetFullPath(filePath);
        byte[] data;
        using (var stream = File.OpenRead(fullPath)) data = OfficeProvenanceBinary.ReadBounded(stream, options.MaxAssetBytes, options.CancellationToken);
        options.CancellationToken.ThrowIfCancellationRequested();
        validatePackage(data, options);
        return OfficeProvenanceInspector.Inspect(data, fullPath, options);
    }

    internal static OfficeProvenanceRemovalResult RemoveFile(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options,
        Func<byte[], OfficeProvenanceRemovalOptions, OfficeProvenanceSignatureStripResult> stripSignatures,
        Func<byte[], OfficeProvenanceRemovalOptions, bool>? hasSignatures = null,
        Action<byte[], OfficeProvenanceOptions>? validatePackage = null,
        bool removeOpcManifestReferences = true,
        bool validateOpcMetadata = true,
        Func<string, bool>? shouldReplacePackageMetadata = null,
        Func<string, byte[], bool, byte[]>? replacePackageMetadata = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeProvenanceRemovalOptions();
        string fullInputPath = Path.GetFullPath(inputPath);
        byte[] data;
        using (var stream = File.OpenRead(fullInputPath)) data = OfficeProvenanceBinary.ReadBounded(stream, options.Limits.MaxAssetBytes, options.Limits.CancellationToken);
        options.Limits.CancellationToken.ThrowIfCancellationRequested();
        OfficeProvenanceRemovalResult result = Remove(
            data,
            fullInputPath,
            options,
            stripSignatures,
            hasSignatures,
            validatePackage,
            removeOpcManifestReferences,
            validateOpcMetadata,
            shouldReplacePackageMetadata,
            replacePackageMetadata);
        OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), result.ToArray());
        return result;
    }

    internal static OfficeProvenanceRemovalResult Remove(
        byte[] data,
        string fileName,
        OfficeProvenanceRemovalOptions? options,
        Func<byte[], OfficeProvenanceRemovalOptions, OfficeProvenanceSignatureStripResult> stripSignatures,
        Func<byte[], OfficeProvenanceRemovalOptions, bool>? hasSignatures = null,
        Action<byte[], OfficeProvenanceOptions>? validatePackage = null,
        bool removeOpcManifestReferences = true,
        bool validateOpcMetadata = true,
        Func<string, bool>? shouldReplacePackageMetadata = null,
        Func<string, byte[], bool, byte[]>? replacePackageMetadata = null) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (stripSignatures == null) throw new ArgumentNullException(nameof(stripSignatures));
        options ??= new OfficeProvenanceRemovalOptions();
        OfficeProvenanceBinary.ValidateRemovalOptions(options);
        options.Limits.CancellationToken.ThrowIfCancellationRequested();
        if (data.LongLength > options.Limits.MaxAssetBytes) {
            throw new InvalidDataException("The package exceeds the configured asset limit.");
        }
        validatePackage?.Invoke(data, options.Limits);
        if (options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.PreserveSignatureMarkup) {
            return OfficeProvenanceRemover.RemoveZipPackage(
                data,
                fileName,
                options,
                removeOpcManifestReferences,
                shouldReplacePackageMetadata,
                replacePackageMetadata);
        }

        OfficeProvenanceRemovalOptions previewOptions = Clone(
            options,
            OfficeSignatureMutationPolicy.PreserveSignatureMarkup,
            options.EffectiveMaxIntermediateBytes,
            options.Limits.MaxExpandedContainerBytes);
        OfficeProvenanceRemovalResult preview = OfficeProvenanceRemover.RemoveZipPackage(
            data,
            fileName,
            previewOptions,
            removeOpcManifestReferences,
            shouldReplacePackageMetadata,
            replacePackageMetadata);
        if (!preview.WasChanged) return EnforceFinalOutputLimit(preview, options.EffectiveMaxOutputBytes);

        OfficeProvenanceZip.ValidateForOwningPackageMutation(data, options.Limits, validateOpcMetadata);
        bool hadSignatureEvidence = hasSignatures?.Invoke(data, options) ?? OfficeProvenanceZip.HasPackageSignature(data, options);
        if (options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.BlockSave) {
            if (hadSignatureEvidence) {
                throw new InvalidOperationException("Removing provenance would invalidate package signatures. Choose an explicit signature mutation policy.");
            }
            return EnforceFinalOutputLimit(preview, options.EffectiveMaxOutputBytes);
        }
        if (!hadSignatureEvidence) return EnforceFinalOutputLimit(preview, options.EffectiveMaxOutputBytes);

        byte[] previewData = preview.ToArray();
        long remainingExpandedBytes = ValidateAggregateRewriteBudget(data, previewData, options.Limits);
        if (remainingExpandedBytes <= 0) {
            throw OfficeProvenanceLimitException.Create("Package mutation exceeds the configured aggregate expanded-byte limit.");
        }
        OfficeProvenanceRemovalOptions stripOptions = Clone(
            options,
            options.SignatureMutationPolicy,
            options.EffectiveMaxOutputBytes,
            remainingExpandedBytes);
        stripOptions.Limits.MaxAssetBytes = Math.Max(
            options.Limits.MaxAssetBytes,
            options.EffectiveMaxOutputBytes);
        OfficeProvenanceSignatureStripResult stripped = stripSignatures(previewData, stripOptions);
        if (!stripped.HadSignatures) {
            throw new InvalidOperationException("The package contains signature evidence that its owning adapter could not remove safely.");
        }
        if (stripped.Data.SequenceEqual(previewData)) {
            throw new InvalidOperationException("The document reports signatures, but its owning package adapter could not remove them safely.");
        }
        if (hasSignatures?.Invoke(stripped.Data, stripOptions) ?? OfficeProvenanceZip.HasPackageSignature(stripped.Data, stripOptions)) {
            throw new InvalidOperationException("The owning package adapter left signature evidence in the rewritten document.");
        }
        OfficeProvenanceBinary.EnsureOutputWithinLimit(stripped.Data.LongLength, options.EffectiveMaxOutputBytes);
        return new OfficeProvenanceRemovalResult(
            stripped.Data,
            preview.Before,
            preview.After,
            preview.Changes,
            wasReserialized: true,
            wereInvalidatedSignaturesRemoved: true);
    }

    private static long ValidateAggregateRewriteBudget(
        byte[] original,
        byte[] preview,
        OfficeProvenanceOptions limits) {
        long expandedBytes = GetExpandedPackageBytes(
            original,
            limits.MaxContainerEntries,
            limits.MaxExpandedContainerBytes,
            limits.CancellationToken);
        long remaining = limits.MaxExpandedContainerBytes - expandedBytes;
        long previewExpandedBytes = GetExpandedPackageBytes(
            preview,
            limits.MaxContainerEntries,
            remaining,
            limits.CancellationToken);
        return remaining - previewExpandedBytes;
    }

    private static long GetExpandedPackageBytes(
        byte[] data,
        int maximumEntries,
        long maximumBytes,
        System.Threading.CancellationToken cancellationToken) {
        if (maximumBytes < 0) {
            throw OfficeProvenanceLimitException.Create("Package mutation exceeds the configured aggregate expanded-byte limit.");
        }
        OfficeProvenanceZip.ValidateEntryCount(data, maximumEntries);
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        long expandedBytes = 0;
        foreach (ZipArchiveEntry entry in archive.Entries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Length > maximumBytes - expandedBytes) {
                throw OfficeProvenanceLimitException.Create("Package mutation exceeds the configured aggregate expanded-byte limit.");
            }
            expandedBytes += entry.Length;
        }
        return expandedBytes;
    }

    private static OfficeProvenanceRemovalOptions Clone(
        OfficeProvenanceRemovalOptions source,
        OfficeSignatureMutationPolicy signaturePolicy,
        long maximumOutputBytes,
        long maximumExpandedBytes) {
        var clone = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = source.RemoveC2paManifests,
            RemoveExternalC2paReferences = source.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = source.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = source.RequireStructurallyValidCarrier,
            SignatureMutationPolicy = signaturePolicy,
            ProcessEmbeddedAssets = source.ProcessEmbeddedAssets && source.Limits.ProcessEmbeddedAssets,
            MaxEmbeddedAssets = Math.Min(source.MaxEmbeddedAssets, source.Limits.MaxEmbeddedAssets),
            MaxOutputBytes = maximumOutputBytes
        };
        clone.Limits.MaxAssetBytes = source.Limits.MaxAssetBytes;
        clone.Limits.MaxManifestBytes = source.Limits.MaxManifestBytes;
        clone.Limits.MaxCarriers = source.Limits.MaxCarriers;
        clone.Limits.MaxContainerEntries = source.Limits.MaxContainerEntries;
        clone.Limits.MaxExpandedContainerBytes = maximumExpandedBytes;
        clone.Limits.CancellationToken = source.Limits.CancellationToken;
        clone.Limits.ProcessEmbeddedAssets = source.ProcessEmbeddedAssets && source.Limits.ProcessEmbeddedAssets;
        clone.Limits.MaxEmbeddedAssets = Math.Min(source.MaxEmbeddedAssets, source.Limits.MaxEmbeddedAssets);
        return clone;
    }

    private static OfficeProvenanceRemovalResult EnforceFinalOutputLimit(
        OfficeProvenanceRemovalResult result,
        long maximumOutputBytes) {
        OfficeProvenanceBinary.EnsureOutputWithinLimit(result.DataLength, maximumOutputBytes);
        return result;
    }
}

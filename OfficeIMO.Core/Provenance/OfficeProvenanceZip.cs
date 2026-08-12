using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceZip {
    private const string ManifestPath = "META-INF/content_credential.c2pa";

    internal static bool HasSignature(byte[] data) => data.Length >= 4 && data[0] == 0x50 && data[1] == 0x4B &&
        ((data[2] == 0x03 && data[3] == 0x04) || (data[2] == 0x05 && data[3] == 0x06) || (data[2] == 0x07 && data[3] == 0x08));

    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        if (archive.Entries.Count > options.MaxContainerEntries) throw new InvalidDataException("ZIP package exceeds the configured entry limit.");
        int index = 0;
        int embeddedCount = 0;
        long expandedBytes = 0;
        foreach (ZipArchiveEntry entry in archive.Entries) {
            if (entry.FullName.Equals(ManifestPath, StringComparison.Ordinal)) {
                if (entry.Length > options.MaxManifestBytes || entry.Length > int.MaxValue) {
                    throw new InvalidDataException("ZIP provenance manifest exceeds the configured manifest limit.");
                }
                ReserveExpandedBytes(ref expandedBytes, entry.Length, options.MaxExpandedContainerBytes);
                byte[] manifest = ReadEntry(entry, (int)entry.Length);
                bool valid = OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.MaxManifestBytes, out _);
                context.Add(new OfficeProvenanceEvidence(
                    OfficeProvenanceCarrierKind.C2paManifest,
                    $"ZIP/{ManifestPath}[{index++}]",
                    valid,
                    manifest.Length));
                continue;
            }
            if (!options.ProcessEmbeddedAssets || !IsSupportedEmbeddedAsset(entry.FullName)) continue;
            embeddedCount++;
            if (embeddedCount > options.MaxEmbeddedAssets) throw new InvalidDataException("ZIP package exceeds the configured embedded-asset limit.");
            if (entry.Length > options.MaxAssetBytes || entry.Length > int.MaxValue) {
                throw new InvalidDataException("A supported embedded asset exceeds the configured asset limit.");
            }
            ReserveExpandedBytes(ref expandedBytes, entry.Length, options.MaxExpandedContainerBytes);
            byte[] asset = ReadEntry(entry, (int)entry.Length);
            try {
                OfficeProvenanceReport nested = OfficeProvenanceInspector.InspectCore(asset, entry.FullName, CreateNestedOptions(options));
                foreach (OfficeProvenanceEvidence evidence in nested.Evidence) context.Add(PrefixEvidence(entry.FullName, evidence));
                foreach (string diagnostic in nested.Diagnostics) context.Diagnostics.Add($"ZIP/{entry.FullName}: {diagnostic}");
            } catch (InvalidDataException exception) {
                context.Diagnostics.Add($"ZIP/{entry.FullName}: embedded asset was preserved because inspection failed: {exception.Message}");
            }
        }
        if (index > 1) context.Diagnostics.Add("The ZIP package contains multiple C2PA manifest entries.");
    }

    internal static byte[] Remove(
        byte[] data,
        OfficeProvenanceRemovalOptions options,
        List<OfficeProvenanceChange> changes,
        out bool reserialized) {
        reserialized = false;
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata && !options.RemoveExternalC2paReferences) return (byte[])data.Clone();
        using var inputStream = new MemoryStream(data, writable: false);
        using var input = new ZipArchive(inputStream, ZipArchiveMode.Read, leaveOpen: false);
        if (input.Entries.Count > options.Limits.MaxContainerEntries) throw new InvalidDataException("ZIP package exceeds the configured entry limit.");
        var removable = new HashSet<string>(StringComparer.Ordinal);
        var embeddedRewrites = new Dictionary<ZipArchiveEntry, byte[]>();
        int occurrence = 0;
        int embeddedCount = 0;
        long inspectionBytes = 0;
        foreach (ZipArchiveEntry entry in input.Entries) {
            if (entry.FullName.Equals(ManifestPath, StringComparison.Ordinal)) {
                if (entry.Length > options.Limits.MaxManifestBytes || entry.Length > int.MaxValue) throw new InvalidDataException("ZIP provenance manifest exceeds the configured limit.");
                ReserveExpandedBytes(ref inspectionBytes, entry.Length, options.Limits.MaxExpandedContainerBytes);
                byte[] manifest = ReadEntry(entry, (int)entry.Length);
                bool valid = OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.Limits.MaxManifestBytes, out _);
                if (options.RemoveC2paManifests && (valid || !options.RequireStructurallyValidCarrier)) {
                    removable.Add(entry.FullName + "\0" + occurrence);
                    changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, $"ZIP/{ManifestPath}[{occurrence}]", entry.Length));
                }
                occurrence++;
                continue;
            }
            if (!options.ProcessEmbeddedAssets || !IsSupportedEmbeddedAsset(entry.FullName)) continue;
            embeddedCount++;
            if (embeddedCount > options.MaxEmbeddedAssets) throw new InvalidDataException("ZIP package exceeds the configured embedded-asset limit.");
            if (entry.Length > options.Limits.MaxAssetBytes || entry.Length > int.MaxValue) throw new InvalidDataException("A supported embedded asset exceeds the configured asset limit.");
            ReserveExpandedBytes(ref inspectionBytes, entry.Length, options.Limits.MaxExpandedContainerBytes);
            byte[] asset = ReadEntry(entry, (int)entry.Length);
            try {
                OfficeProvenanceRemovalResult nested = OfficeProvenanceRemover.Remove(asset, entry.FullName, CreateNestedRemovalOptions(options));
                if (!nested.WasChanged) continue;
                embeddedRewrites.Add(entry, nested.ToArray());
                foreach (OfficeProvenanceChange change in nested.Changes) {
                    changes.Add(new OfficeProvenanceChange(change.Carrier, $"ZIP/{entry.FullName}/{change.Location}", change.RemovedBytes));
                }
                if (nested.WasReserialized) reserialized = true;
            } catch (InvalidDataException) {
                // Malformed embedded assets are preserved; document-level diagnostics are available during inspection.
            }
        }
        if (changes.Count == 0) return (byte[])data.Clone();

        bool hasSignature = input.Entries.Any(IsSignatureEntry);
        if (hasSignature && options.SignatureMutationPolicy == OfficeIMO.OfficeSignatureMutationPolicy.BlockSave) {
            throw new InvalidOperationException("Removing provenance would invalidate package signatures. Choose an explicit signature mutation policy.");
        }
        if (hasSignature && options.SignatureMutationPolicy == OfficeIMO.OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures) {
            throw new NotSupportedException("Generic ZIP provenance removal cannot safely rewrite package signatures. Use the owning OfficeIMO document format API.");
        }

        using var outputStream = new MemoryStream(data.Length);
        using (var output = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true)) {
            long expandedBytes = 0;
            occurrence = 0;
            IEnumerable<ZipArchiveEntry> orderedEntries = input.Entries
                .OrderByDescending(entry => entry.FullName.Equals("mimetype", StringComparison.Ordinal));
            foreach (ZipArchiveEntry entry in orderedEntries) {
                bool isManifest = entry.FullName.Equals(ManifestPath, StringComparison.Ordinal);
                string key = entry.FullName + "\0" + occurrence;
                if (isManifest) occurrence++;
                if (isManifest && removable.Contains(key)) continue;
                bool isMimetype = entry.FullName.Equals("mimetype", StringComparison.Ordinal);
                ZipArchiveEntry target = output.CreateEntry(entry.FullName, isMimetype ? CompressionLevel.NoCompression : CompressionLevel.Optimal);
                if (!isMimetype) target.LastWriteTime = entry.LastWriteTime;
                using Stream destination = target.Open();
                if (embeddedRewrites.TryGetValue(entry, out byte[]? rewritten)) {
                    ReserveExpandedBytes(ref expandedBytes, rewritten.LongLength, options.Limits.MaxExpandedContainerBytes);
                    destination.Write(rewritten, 0, rewritten.Length);
                } else {
                    using Stream source = entry.Open();
                    CopyBounded(source, destination, options.Limits.MaxAssetBytes, options.Limits.MaxExpandedContainerBytes, ref expandedBytes);
                }
            }
        }
        reserialized = true;
        return outputStream.ToArray();
    }

    private static bool IsSignatureEntry(ZipArchiveEntry entry) =>
        entry.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase) ||
        entry.FullName.Equals("META-INF/documentsignatures.xml", StringComparison.OrdinalIgnoreCase) ||
        entry.FullName.Equals("META-INF/macrosignatures.xml", StringComparison.OrdinalIgnoreCase) ||
        entry.FullName.Equals("META-INF/signatures.xml", StringComparison.OrdinalIgnoreCase);

    private static bool IsSupportedEmbeddedAsset(string name) {
        string extension = Path.GetExtension(name).ToLowerInvariant();
        return extension is ".jpg" or ".jpeg" or ".png" or ".webp" or ".gif" or ".tif" or ".tiff" or ".svg";
    }

    private static OfficeProvenanceOptions CreateNestedOptions(OfficeProvenanceOptions source) => new OfficeProvenanceOptions {
        MaxAssetBytes = source.MaxAssetBytes,
        MaxManifestBytes = source.MaxManifestBytes,
        MaxCarriers = source.MaxCarriers,
        MaxContainerEntries = source.MaxContainerEntries,
        MaxExpandedContainerBytes = source.MaxExpandedContainerBytes,
        ProcessEmbeddedAssets = false,
        MaxEmbeddedAssets = source.MaxEmbeddedAssets
    };

    private static OfficeProvenanceRemovalOptions CreateNestedRemovalOptions(OfficeProvenanceRemovalOptions source) {
        var nested = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = source.RemoveC2paManifests,
            RemoveExternalC2paReferences = source.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = source.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = source.RequireStructurallyValidCarrier,
            SignatureMutationPolicy = source.SignatureMutationPolicy,
            ProcessEmbeddedAssets = false,
            MaxEmbeddedAssets = source.MaxEmbeddedAssets
        };
        nested.Limits.MaxAssetBytes = source.Limits.MaxAssetBytes;
        nested.Limits.MaxManifestBytes = source.Limits.MaxManifestBytes;
        nested.Limits.MaxCarriers = source.Limits.MaxCarriers;
        nested.Limits.MaxContainerEntries = source.Limits.MaxContainerEntries;
        nested.Limits.MaxExpandedContainerBytes = source.Limits.MaxExpandedContainerBytes;
        nested.Limits.ProcessEmbeddedAssets = false;
        nested.Limits.MaxEmbeddedAssets = source.Limits.MaxEmbeddedAssets;
        return nested;
    }

    private static OfficeProvenanceEvidence PrefixEvidence(string entryName, OfficeProvenanceEvidence evidence) =>
        new OfficeProvenanceEvidence(
            evidence.Carrier,
            $"ZIP/{entryName}/{evidence.Location}",
            evidence.IsStructurallyValid,
            evidence.PayloadLength,
            evidence.Value,
            evidence.DigitalSourceKind);

    private static byte[] ReadEntry(ZipArchiveEntry entry, int length) {
        byte[] data = new byte[length];
        using Stream stream = entry.Open();
        OfficeProvenanceBinary.ReadExactly(stream, data, 0, data.Length);
        if (stream.ReadByte() != -1) throw new InvalidDataException("ZIP entry expanded beyond its declared length.");
        return data;
    }

    private static void CopyBounded(Stream source, Stream destination, long maximumEntryBytes, long maximumTotalBytes, ref long totalBytes) {
        byte[] buffer = new byte[81920];
        long entryBytes = 0;
        while (true) {
            int read = source.Read(buffer, 0, buffer.Length);
            if (read <= 0) break;
            if (entryBytes > maximumEntryBytes - read) throw new InvalidDataException("ZIP entry exceeds the configured copy limit.");
            entryBytes += read;
            ReserveExpandedBytes(ref totalBytes, read, maximumTotalBytes);
            destination.Write(buffer, 0, read);
        }
    }

    private static void ReserveExpandedBytes(ref long totalBytes, long bytes, long maximumBytes) {
        if (bytes < 0 || totalBytes > maximumBytes - bytes) throw new InvalidDataException("ZIP package exceeds the configured expanded-byte limit.");
        totalBytes += bytes;
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Options controlling how read-result assets are materialized outside the read result envelope.
/// </summary>
public sealed class OfficeDocumentAssetMaterializationOptions {
    /// <summary>
    /// Creates the destination directory when it does not exist. Defaults to true.
    /// </summary>
    public bool CreateDirectory { get; set; } = true;

    /// <summary>
    /// Overwrites an existing file with the same deterministic filename. Defaults to true.
    /// </summary>
    public bool Overwrite { get; set; } = true;

    /// <summary>
    /// When true, validates an asset payload against <see cref="OfficeDocumentAsset.PayloadHash"/> when the hash is present.
    /// </summary>
    public bool ValidatePayloadHash { get; set; }

    /// <summary>
    /// Optional asset predicate used to materialize only selected assets.
    /// </summary>
    public Func<OfficeDocumentAsset, bool>? Predicate { get; set; }
}

/// <summary>
/// Result for a single asset materialization attempt.
/// </summary>
public sealed class OfficeDocumentMaterializedAsset {
    /// <summary>
    /// Asset from the source read result.
    /// </summary>
    public OfficeDocumentAsset Asset { get; set; } = new OfficeDocumentAsset();

    /// <summary>
    /// Deterministic filename used for this materialization attempt.
    /// </summary>
    public string? FileName { get; set; }

    /// <summary>
    /// Full output path when assets are written to a directory.
    /// </summary>
    public string? Path { get; set; }

    /// <summary>
    /// True when the asset payload was written or streamed.
    /// </summary>
    public bool Written { get; set; }

    /// <summary>
    /// Explanation when the asset was skipped.
    /// </summary>
    public string? SkippedReason { get; set; }
}

/// <summary>
/// Helpers for writing or streaming materializable assets from a shared read result.
/// </summary>
public static class OfficeDocumentAssetMaterializer {
    /// <summary>
    /// Writes materializable asset payloads to <paramref name="directoryPath"/> using each asset's deterministic filename.
    /// </summary>
    /// <param name="result">Read result that owns the assets.</param>
    /// <param name="directoryPath">Destination directory.</param>
    /// <param name="options">Materialization options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<OfficeDocumentMaterializedAsset> WriteAssetsToDirectory(
        this OfficeDocumentReadResult result,
        string directoryPath,
        OfficeDocumentAssetMaterializationOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (string.IsNullOrWhiteSpace(directoryPath)) throw new ArgumentException("Directory path cannot be empty.", nameof(directoryPath));

        OfficeDocumentAssetMaterializationOptions effectiveOptions = options ?? new OfficeDocumentAssetMaterializationOptions();
        if (effectiveOptions.CreateDirectory) {
            Directory.CreateDirectory(directoryPath);
        } else if (!Directory.Exists(directoryPath)) {
            throw new DirectoryNotFoundException("Directory '" + directoryPath + "' does not exist.");
        }

        var results = new List<OfficeDocumentMaterializedAsset>();
        foreach (OfficeDocumentAsset asset in SelectAssets(result, effectiveOptions)) {
            cancellationToken.ThrowIfCancellationRequested();
            string fileName = ResolveFileName(asset);
            string outputPath = System.IO.Path.Combine(directoryPath, fileName);
            byte[]? payload = asset.PayloadBytes;

            if (payload == null || payload.Length == 0) {
                results.Add(Skipped(asset, fileName, outputPath, "Asset has no in-memory payload."));
                continue;
            }

            if (HasPayloadHashMismatch(asset, effectiveOptions)) {
                results.Add(Skipped(asset, fileName, outputPath, "Asset payload hash does not match PayloadHash."));
                continue;
            }

            if (!effectiveOptions.Overwrite && File.Exists(outputPath)) {
                results.Add(Skipped(asset, fileName, outputPath, "Destination file already exists."));
                continue;
            }

            ReaderFileCommit.WriteAllBytes(outputPath, payload);
            results.Add(new OfficeDocumentMaterializedAsset {
                Asset = asset,
                FileName = fileName,
                Path = outputPath,
                Written = true
            });
        }

        return results;
    }

    /// <summary>
    /// Streams materializable asset payloads to a caller-owned callback without writing files.
    /// </summary>
    /// <param name="result">Read result that owns the assets.</param>
    /// <param name="writeAsset">Callback that receives each asset and a readable payload stream.</param>
    /// <param name="options">Materialization options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<OfficeDocumentMaterializedAsset> StreamAssets(
        this OfficeDocumentReadResult result,
        Action<OfficeDocumentAsset, Stream> writeAsset,
        OfficeDocumentAssetMaterializationOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        if (writeAsset == null) throw new ArgumentNullException(nameof(writeAsset));

        OfficeDocumentAssetMaterializationOptions effectiveOptions = options ?? new OfficeDocumentAssetMaterializationOptions();
        var results = new List<OfficeDocumentMaterializedAsset>();
        foreach (OfficeDocumentAsset asset in SelectAssets(result, effectiveOptions)) {
            cancellationToken.ThrowIfCancellationRequested();
            string fileName = ResolveFileName(asset);
            byte[]? payload = asset.PayloadBytes;

            if (payload == null || payload.Length == 0) {
                results.Add(Skipped(asset, fileName, null, "Asset has no in-memory payload."));
                continue;
            }

            if (HasPayloadHashMismatch(asset, effectiveOptions)) {
                results.Add(Skipped(asset, fileName, null, "Asset payload hash does not match PayloadHash."));
                continue;
            }

            using var stream = new MemoryStream(payload, writable: false);
            writeAsset(asset, stream);
            results.Add(new OfficeDocumentMaterializedAsset {
                Asset = asset,
                FileName = fileName,
                Written = true
            });
        }

        return results;
    }

    private static IEnumerable<OfficeDocumentAsset> SelectAssets(OfficeDocumentReadResult result, OfficeDocumentAssetMaterializationOptions options) {
        IReadOnlyList<OfficeDocumentAsset> assets = result.Assets ?? Array.Empty<OfficeDocumentAsset>();
        return options.Predicate == null ? assets : assets.Where(options.Predicate);
    }

    private static string ResolveFileName(OfficeDocumentAsset asset) {
        string fileName = string.IsNullOrWhiteSpace(asset.FileName)
            ? OfficeDocumentAssetNaming.BuildFileName(asset.Id, asset.Extension)
            : asset.FileName!;
        fileName = System.IO.Path.GetFileName(fileName);
        return string.IsNullOrWhiteSpace(fileName) ? OfficeDocumentAssetNaming.BuildFileName(asset.Id, asset.Extension) : fileName;
    }

    private static bool HasPayloadHashMismatch(OfficeDocumentAsset asset, OfficeDocumentAssetMaterializationOptions options) {
        if (!options.ValidatePayloadHash || string.IsNullOrWhiteSpace(asset.PayloadHash)) {
            return false;
        }

        return !asset.PayloadHashMatches(out _);
    }

    private static OfficeDocumentMaterializedAsset Skipped(OfficeDocumentAsset asset, string fileName, string? path, string reason) {
        return new OfficeDocumentMaterializedAsset {
            Asset = asset,
            FileName = fileName,
            Path = path,
            Written = false,
            SkippedReason = reason
        };
    }
}

using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

/// <summary>
/// Options controlling opt-in data URI generation for materializable assets.
/// </summary>
public sealed class OfficeDocumentAssetDataUriOptions {
    /// <summary>
    /// Maximum payload size to inline. Defaults to 64 KiB. Set to null for no size cap.
    /// </summary>
    public long? MaxInlineBytes { get; set; } = 64L * 1024L;

    /// <summary>
    /// Optional asset predicate used to inline only selected assets.
    /// </summary>
    public Func<OfficeDocumentAsset, bool>? Predicate { get; set; }
}

/// <summary>
/// Opt-in data URI helpers for materializable read-result assets.
/// </summary>
public static class OfficeDocumentAssetDataUri {
    /// <summary>
    /// Builds a data URI for an asset payload when the payload exists and fits the configured size cap.
    /// </summary>
    /// <param name="asset">Asset to encode.</param>
    /// <param name="dataUri">Data URI when one could be built.</param>
    /// <param name="maxInlineBytes">Maximum payload size to inline. Set to null for no size cap.</param>
    public static bool TryBuildDataUri(this OfficeDocumentAsset asset, out string? dataUri, long? maxInlineBytes = 64L * 1024L) {
        if (asset == null) throw new ArgumentNullException(nameof(asset));
        if (maxInlineBytes.HasValue && maxInlineBytes.Value < 0) throw new ArgumentOutOfRangeException(nameof(maxInlineBytes), "Maximum inline bytes cannot be negative.");

        byte[]? payload = asset.PayloadBytes;
        if (payload == null || payload.Length == 0) {
            dataUri = null;
            return false;
        }

        if (maxInlineBytes.HasValue && payload.LongLength > maxInlineBytes.Value) {
            dataUri = null;
            return false;
        }

        string mediaType = string.IsNullOrWhiteSpace(asset.MediaType) ? "application/octet-stream" : asset.MediaType!.Trim();
        dataUri = "data:" + mediaType + ";base64," + Convert.ToBase64String(payload);
        return true;
    }

    /// <summary>
    /// Builds a data URI map keyed by asset id for materializable assets in a read result.
    /// </summary>
    /// <param name="result">Read result that owns the assets.</param>
    /// <param name="options">Data URI options.</param>
    public static IReadOnlyDictionary<string, string> BuildAssetDataUriMap(
        this OfficeDocumentReadResult result,
        OfficeDocumentAssetDataUriOptions? options = null) {
        if (result == null) throw new ArgumentNullException(nameof(result));

        OfficeDocumentAssetDataUriOptions effectiveOptions = options ?? new OfficeDocumentAssetDataUriOptions();
        if (effectiveOptions.MaxInlineBytes.HasValue && effectiveOptions.MaxInlineBytes.Value < 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "Maximum inline bytes cannot be negative.");
        }

        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        IReadOnlyList<OfficeDocumentAsset> assets = result.Assets ?? Array.Empty<OfficeDocumentAsset>();
        foreach (OfficeDocumentAsset asset in assets) {
            if (string.IsNullOrWhiteSpace(asset.Id)) {
                continue;
            }

            if (effectiveOptions.Predicate != null && !effectiveOptions.Predicate(asset)) {
                continue;
            }

            if (asset.TryBuildDataUri(out string? dataUri, effectiveOptions.MaxInlineBytes) && dataUri != null) {
                map[asset.Id] = dataUri;
            }
        }

        return map;
    }
}

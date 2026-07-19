using System;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Reader;

/// <summary>
/// Hash helpers for materializable read-result assets.
/// </summary>
public static class OfficeDocumentAssetHash {
    /// <summary>
    /// Computes a lowercase SHA-256 hex hash for asset payload bytes.
    /// </summary>
    /// <param name="payload">Asset payload bytes.</param>
    public static string ComputeSha256Hex(byte[] payload) {
        if (payload == null) throw new ArgumentNullException(nameof(payload));

        using var sha = SHA256.Create();
        byte[] hash = sha.ComputeHash(payload);
        var builder = new StringBuilder(hash.Length * 2);
        for (int i = 0; i < hash.Length; i++) {
            builder.Append(hash[i].ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    /// <summary>
    /// Checks whether an asset payload matches its declared payload hash.
    /// </summary>
    /// <param name="asset">Asset to validate.</param>
    /// <param name="actualHash">Computed SHA-256 hash when validation could run.</param>
    public static bool PayloadHashMatches(this OfficeDocumentAsset asset, out string? actualHash) {
        if (asset == null) throw new ArgumentNullException(nameof(asset));

        byte[]? payload = asset.PayloadBytes;
        if (payload == null || payload.Length == 0 || string.IsNullOrWhiteSpace(asset.PayloadHash)) {
            actualHash = null;
            return false;
        }

        string expectedHash = asset.PayloadHash!.Trim();
        actualHash = ComputeSha256Hex(payload);
        return string.Equals(expectedHash, actualHash, StringComparison.OrdinalIgnoreCase);
    }
}

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Hash helpers for materializable read-result assets.
/// </summary>
public static class OfficeDocumentAssetHash {
    /// <summary>Hashes the remaining stream bytes while observing cancellation. Leaves the stream open at the consumed position.</summary>
    public static string ComputeSha256Hex(Stream stream, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        using var sha = SHA256.Create();
        var buffer = new byte[81920];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = stream.Read(buffer, 0, buffer.Length);
            cancellationToken.ThrowIfCancellationRequested();
            if (count == 0) break;
            sha.TransformBlock(buffer, 0, count, buffer, 0);
        }
        sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        var builder = new StringBuilder(64);
        foreach (byte value in sha.Hash!) builder.Append(value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
        return builder.ToString();
    }

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

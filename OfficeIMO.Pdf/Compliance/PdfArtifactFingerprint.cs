using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

internal static class PdfArtifactFingerprint {
    internal static string ComputeSha256(byte[] artifact) {
        Guard.NotNull(artifact, nameof(artifact));
        byte[] hash;
#if NET6_0_OR_GREATER
        hash = SHA256.HashData(artifact);
#else
        using (SHA256 sha256 = SHA256.Create()) {
            hash = sha256.ComputeHash(artifact);
        }
#endif
        {
            var builder = new System.Text.StringBuilder(hash.Length * 2);
            for (int i = 0; i < hash.Length; i++) {
                builder.Append(hash[i].ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }
    }

    internal static string? NormalizeSha256(string? value, string paramName) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        string normalized = value!.Trim().ToLowerInvariant();
        if (normalized.Length != 64 || normalized.Any(static ch =>
                !char.IsDigit(ch) && (ch < 'a' || ch > 'f'))) {
            throw new System.ArgumentException("Artifact SHA-256 must contain exactly 64 hexadecimal characters.", paramName);
        }

        return normalized;
    }
}

namespace OfficeIMO.Pdf;

internal static class PdfArtifactFingerprint {
    internal static string ComputeSha256(byte[] artifact) {
        Guard.NotNull(artifact, nameof(artifact));
        return ComputeSha256((ReadOnlyMemory<byte>)artifact);
    }

    internal static string ComputeSha256(ReadOnlyMemory<byte> artifact) =>
        PdfArtifactInfo.ComputeSha256Hex(artifact, CancellationToken.None);

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

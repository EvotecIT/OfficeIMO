using System.Text;

namespace OfficeIMO.Pdf;

internal static class PdfModernPasswordNormalizer {
    internal const int MaximumInputCharacters = 4096;

    /// <summary>
    /// Prepares a revision 5+ password for writing: SASLprep, UTF-8, at most 127 bytes. A password SASLprep prohibits
    /// (control, private-use or unpaired surrogate characters) keeps its NFKC form, so existing output is unchanged.
    /// </summary>
    internal static byte[] Normalize(string password, CancellationToken cancellationToken) =>
        GetAuthenticationCandidates(password, cancellationToken)[0];

    /// <summary>
    /// Password byte candidates for authentication. The SASLprep form comes first; the plain NFKC form follows when it
    /// differs, because some producers (including OfficeIMO before SASLprep) skipped the stringprep mapping.
    /// </summary>
    internal static IReadOnlyList<byte[]> GetAuthenticationCandidates(string password, CancellationToken cancellationToken) {
        ValidateLength(password, cancellationToken);
        var candidates = new List<byte[]>(2);
        if (PdfSaslPrep.TryPrepare(password, out string prepared)) {
            candidates.Add(ToPasswordBytes(prepared));
        }

        cancellationToken.ThrowIfCancellationRequested();
        string normalized;
        try {
            normalized = password.Normalize(NormalizationForm.FormKC);
        } catch (ArgumentException) {
            normalized = password;
        }

        byte[] legacy = ToPasswordBytes(normalized);
        if (candidates.Count == 0 || !candidates[0].SequenceEqual(legacy)) {
            candidates.Add(legacy);
        }

        return candidates;
    }

    private static void ValidateLength(string password, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        // Revisions 5 and later use at most 127 UTF-8 bytes. Bound the caller input before
        // normalization can scan and allocate an arbitrarily large string.
        if (password.Length > MaximumInputCharacters) {
            throw new ArgumentOutOfRangeException(nameof(password),
                "PDF Standard revision 5 and later passwords cannot exceed 4096 characters before normalization.");
        }
    }

    private static byte[] ToPasswordBytes(string value) {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        if (bytes.Length <= 127) return bytes;

        var truncated = new byte[127];
        Buffer.BlockCopy(bytes, 0, truncated, 0, truncated.Length);
        return truncated;
    }
}

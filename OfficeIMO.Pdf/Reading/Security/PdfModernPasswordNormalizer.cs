using System.Text;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfModernPasswordNormalizer {
    internal const int MaximumInputCharacters = 4096;

    internal static byte[] Normalize(string password, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        // Revisions 5 and 6 use at most 127 UTF-8 bytes. Bound the caller input before
        // FormKC normalization can scan and allocate an arbitrarily large string.
        if (password.Length > MaximumInputCharacters) {
            throw new ArgumentOutOfRangeException(nameof(password),
                "PDF Standard revision 5/6 passwords cannot exceed 4096 characters before normalization.");
        }

        string normalized = password.Normalize(NormalizationForm.FormKC);
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = Encoding.UTF8.GetBytes(normalized);
        cancellationToken.ThrowIfCancellationRequested();
        if (bytes.Length <= 127) return bytes;

        var truncated = new byte[127];
        Buffer.BlockCopy(bytes, 0, truncated, 0, truncated.Length);
        return truncated;
    }
}

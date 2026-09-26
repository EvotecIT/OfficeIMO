using System.Text;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfLegacyPasswordEncoding {
    internal static byte[] EncodePrefix(string password, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrEmpty(password)) return Array.Empty<byte>();

        // Standard revisions 2-4 consume only the first 32 encoded bytes. The
        // encoding choice still considers the whole password, with cancellation
        // polling, but materialization needs only a bounded prefix. A 64-unit UTF-16
        // prefix covers the first 32 UTF-8 bytes and any crossing surrogate pair.
        bool useWinAnsi = PdfWinAnsiEncoding.CanEncode(password, out _, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        string prefix = password.Substring(0, Math.Min(password.Length, useWinAnsi ? 32 : 64));
        byte[] bytes = useWinAnsi
            ? PdfWinAnsiEncoding.Encode(prefix, cancellationToken)
            : Encoding.UTF8.GetBytes(prefix);
        cancellationToken.ThrowIfCancellationRequested();
        return bytes;
    }
}

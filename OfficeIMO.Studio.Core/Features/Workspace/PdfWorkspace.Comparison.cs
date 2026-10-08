using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    /// <summary>Publishes a bounded gallery only while both captured sources and the edited revision remain current.</summary>
    internal async Task ExportComparisonReportAsync(string destination, PdfVisualComparisonReport report,
        long expectedRevision, string comparisonPath, CancellationToken cancellationToken,
        Func<bool>? isCurrent = null) {
        ThrowIfDisposed();
        if (!System.IO.Path.GetExtension(_storage.Describe(destination).Name).Equals(".html", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Comparison output must be an HTML gallery.", nameof(destination));
        async Task VerifyAsync(CancellationToken token) {
            token.ThrowIfCancellationRequested();
            if (_revision != expectedRevision || isCurrent?.Invoke() == false ||
                !string.Equals(PdfWorkspaceRecoveryStore.Fingerprint(_bytes), report.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new IOException("The document or comparison scope changed. Compare again before exporting.");
            if (OfficeStorageIdentity.AreEquivalent(destination, comparisonPath))
                throw new IOException("The comparison report cannot replace either source PDF.");
            if (!string.Equals(await _storage.FingerprintAsync(Path, token).ConfigureAwait(false), _baseFingerprint, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(await _storage.FingerprintAsync(comparisonPath, token).ConfigureAwait(false), report.ActualSha256, StringComparison.OrdinalIgnoreCase))
                throw new IOException("A source PDF changed after comparison. Open it again and compare before exporting.");
        }
        string gallery = report.ToHtmlGallery("OfficeIMO document comparison", 96L * 1024 * 1024, cancellationToken);
        byte[] bytes = new UTF8Encoding(false).GetBytes(gallery);
        await WriteWorkspaceOutputAsync(destination, (stream, token) => stream.WriteAsync(bytes.AsMemory(), token).AsTask(),
            cancellationToken, VerifyAsync).ConfigureAwait(false);
    }
}

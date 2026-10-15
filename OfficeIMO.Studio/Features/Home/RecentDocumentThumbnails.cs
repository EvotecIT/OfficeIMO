using Avalonia.Media.Imaging;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Home;

/// <summary>
/// Renders small first-page previews for local recent documents. Each Home card owns its preview and cancellation;
/// private page images are not retained in a process-wide cache after the card clears them.
/// </summary>
internal static class RecentDocumentThumbnails {
    internal const double TargetWidth = 176D;
    private static readonly SemaphoreSlim Gate = new(2, 2);

    internal static async Task<RecentDocumentPreview?> GetAsync(string path, CancellationToken cancellationToken = default) {
        string? key = GetFingerprint(path);
        if (key is null) return null;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            string local = OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(path)!;
            return await RenderAsync(local, cancellationToken).ConfigureAwait(false);
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) {
            return null;
        }
    }

    internal static string? GetFingerprint(string path) {
        try {
            if (OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(path) is not { } local ||
                !string.Equals(System.IO.Path.GetExtension(local), ".pdf", StringComparison.OrdinalIgnoreCase)) return null;
            var file = new FileInfo(local);
            return file.Exists ? $"{file.FullName}|{file.Length}|{file.LastWriteTimeUtc.Ticks}" : null;
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) {
            return null;
        }
    }

    private static async Task<RecentDocumentPreview?> RenderAsync(string path, CancellationToken cancellationToken) {
        await Gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(TimeSpan.FromSeconds(8));
            PdfDocumentSession session = await PdfDocumentSession.OpenAsync(path, timeout.Token).ConfigureAwait(false);
            if (session.Pages.Count == 0) return new RecentDocumentPreview(null, 0);
            PdfPageInfo first = session.Pages[0];
            bool swapsAxes = Math.Abs(first.RotationDegrees) % 180 == 90;
            double width = Math.Max(1D, swapsAxes ? first.Height : first.Width);
            double scale = Math.Clamp(TargetWidth * 1.5D / width, 0.05D, 1D);
            PdfRenderedPage rendered = await session.RenderPageAsync(1, scale, timeout.Token).ConfigureAwait(false);
            timeout.Token.ThrowIfCancellationRequested();
            using var stream = new MemoryStream(rendered.Bytes, writable: false);
            return new RecentDocumentPreview(new Bitmap(stream), session.Pages.Count);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception error) when (error is not OutOfMemoryException) {
            // Protected, damaged, or slow documents keep the generic document card.
            return null;
        } finally {
            Gate.Release();
        }
    }
}

internal sealed record RecentDocumentPreview(Bitmap? Image, int PageCount);

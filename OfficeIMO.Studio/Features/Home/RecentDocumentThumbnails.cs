using System.Collections.Concurrent;
using Avalonia.Media.Imaging;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Home;

/// <summary>
/// Renders small first-page previews for local recent documents. Previews stay in memory for the
/// process lifetime, keyed by path, size and write time so an edited file gets a fresh preview.
/// </summary>
internal static class RecentDocumentThumbnails {
    internal const double TargetWidth = 176D;
    private const int MaximumEntries = 32;
    private static readonly ConcurrentDictionary<string, Lazy<Task<RecentDocumentPreview?>>> Cache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly SemaphoreSlim Gate = new(2, 2);

    internal static Task<RecentDocumentPreview?> GetAsync(string path) {
        if (OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(path) is not { } local ||
            !string.Equals(System.IO.Path.GetExtension(local), ".pdf", StringComparison.OrdinalIgnoreCase)) {
            return Task.FromResult<RecentDocumentPreview?>(null);
        }
        FileInfo file;
        try {
            file = new FileInfo(local);
            if (!file.Exists) return Task.FromResult<RecentDocumentPreview?>(null);
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) {
            return Task.FromResult<RecentDocumentPreview?>(null);
        }
        string key = $"{file.FullName}|{file.Length}|{file.LastWriteTimeUtc.Ticks}";
        if (Cache.Count > MaximumEntries) Cache.Clear();
        return Cache.GetOrAdd(key, _ => new Lazy<Task<RecentDocumentPreview?>>(() => RenderAsync(file.FullName))).Value;
    }

    private static async Task<RecentDocumentPreview?> RenderAsync(string path) {
        await Gate.WaitAsync().ConfigureAwait(false);
        try {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(8));
            PdfDocumentSession session = await PdfDocumentSession.OpenAsync(path, timeout.Token).ConfigureAwait(false);
            if (session.Pages.Count == 0) return new RecentDocumentPreview(null, 0);
            PdfPageInfo first = session.Pages[0];
            bool swapsAxes = Math.Abs(first.RotationDegrees) % 180 == 90;
            double width = Math.Max(1D, swapsAxes ? first.Height : first.Width);
            double scale = Math.Clamp(TargetWidth * 1.5D / width, 0.05D, 1D);
            PdfRenderedPage rendered = await session.RenderPageAsync(1, scale, timeout.Token).ConfigureAwait(false);
            using var stream = new MemoryStream(rendered.Bytes, writable: false);
            return new RecentDocumentPreview(new Bitmap(stream), session.Pages.Count);
        } catch (Exception error) when (error is not OutOfMemoryException) {
            // Protected, damaged, or slow documents keep the generic document card.
            return null;
        } finally {
            Gate.Release();
        }
    }
}

internal sealed record RecentDocumentPreview(Bitmap? Image, int PageCount);

using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Reader;

internal static class StudioPdfSecurityPolicy {
    internal const int MaximumPages = 1_000;
    internal const long MaximumInputBytes = 64L * 1024L * 1024L;
    // Letter and A4 pages fit at 4x, so 2x-3x displays keep device resolution through normal zoom levels.
    internal const long MaximumRasterPixels = 8_400_000;
    // Uncompressed RGBA plus PNG framing, so the pixel budget rather than encoding governs a photographic page.
    internal const long MaximumRasterOutputBytes = MaximumRasterPixels * 4L + 1024L * 1024L;
    internal static readonly TimeSpan RenderTimeout = TimeSpan.FromSeconds(10);

    internal static PdfLoadOptions CreateLoadOptions(string? password = null) => new() {
        Password = password,
        Limits = new PdfReadLimits {
            MaxInputBytes = MaximumInputBytes,
            MaxIndirectObjects = 50_000,
            MaxRawStreamBytes = 16 * 1024 * 1024,
            MaxDecodedStreamBytes = 16 * 1024 * 1024,
            MaxTotalDecodedStreamBytes = 64L * 1024L * 1024L,
            MaxPageContentBytes = 16 * 1024 * 1024,
            MaxRetainedContentBytes = 64L * 1024L * 1024L,
            MaxPages = MaximumPages,
            MaxObjectParsingTime = TimeSpan.FromSeconds(10)
        }
    };
}

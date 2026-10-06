using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Reader;

internal static class StudioPdfSecurityPolicy {
    internal const int MaximumPages = 1_000;
    internal const long MaximumInputBytes = 64L * 1024L * 1024L;
    internal const long MaximumRasterPixels = 4_000_000;
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

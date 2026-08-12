namespace OfficeIMO.Pdf;

/// <summary>
/// Reuses immutable font decoding state for pages and forms that reference the same resource dictionary.
/// The cache is owned by one parsed document, so entries cannot cross document object graphs.
/// </summary>
internal sealed class PdfFontResourceCache {
    private const int MaxFontResources = 64;
    private const long MaxEstimatedRetainedFontBytes = 64L * 1024 * 1024;
    private static readonly PdfFontResourceSet Empty = new(
        new Dictionary<string, PdfFontResource>(StringComparer.Ordinal),
        new Dictionary<string, Func<byte[], int, string>>(StringComparer.Ordinal),
        new Dictionary<string, Func<byte[], double>>(StringComparer.Ordinal));

    private readonly Dictionary<PdfDictionary, PdfFontResource> _fonts = new();
    private readonly object _sync = new();
    private long _estimatedRetainedFontBytes;

    internal PdfFontResourceSet GetOrCreate(
        PdfDictionary? resources,
        Dictionary<int, PdfIndirectObject> objects) {
        if (resources is null) {
            return Empty;
        }

        lock (_sync) {
            return ResourceResolver.CreateFontResourceSet(resources, objects, GetOrCreateFont);
        }
    }

    private PdfFontResource GetOrCreateFont(
        string resourceName,
        PdfDictionary font,
        Dictionary<int, PdfIndirectObject> objects) {
        if (_fonts.TryGetValue(font, out PdfFontResource? existing)) {
            return existing.WithResourceName(resourceName);
        }

        PdfFontResource created = ResourceResolver.CreateFontResource(resourceName, font, objects);
        long estimatedBytes = EstimateRetainedBytes(created);
        if (_fonts.Count < MaxFontResources &&
            estimatedBytes <= MaxEstimatedRetainedFontBytes - _estimatedRetainedFontBytes) {
            _fonts.Add(font, created);
            _estimatedRetainedFontBytes += estimatedBytes;
        }
        return created;
    }

    private static long EstimateRetainedBytes(PdfFontResource font) {
        long embeddedBytes = font.EmbeddedTrueTypeFont?.LongLength ?? 0;
        long cmapBytes = (font.CMap?.MappingCount ?? 0) * 128L;
        long differencesBytes = (font.Differences?.Count ?? 0) * 64L;
        return embeddedBytes + cmapBytes + differencesBytes + 1024L;
    }
}

internal sealed class PdfFontResourceSet {
    internal PdfFontResourceSet(
        Dictionary<string, PdfFontResource> fonts,
        Dictionary<string, Func<byte[], int, string>> decoders,
        Dictionary<string, Func<byte[], double>> widthProviders) {
        Fonts = fonts;
        Decoders = decoders;
        WidthProviders = widthProviders;
    }

    internal Dictionary<string, PdfFontResource> Fonts { get; }
    internal Dictionary<string, Func<byte[], int, string>> Decoders { get; }
    internal Dictionary<string, Func<byte[], double>> WidthProviders { get; }
}
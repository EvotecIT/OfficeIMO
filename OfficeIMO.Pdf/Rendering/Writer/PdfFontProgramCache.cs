using System.Runtime.CompilerServices;

namespace OfficeIMO.Pdf;

/// <summary>
/// Reuses immutable font parsing data while keeping glyph usage isolated per
/// document. Callers provide private, immutable byte snapshots; byte-array keys
/// are weak and each parsed blueprint owns its own private byte snapshot.
/// </summary>
internal static class PdfFontProgramCache {
    private static readonly ConditionalWeakTable<byte[], CacheBucket> Cache = new();

    internal static PdfTrueTypeFontProgram GetTrueType(byte[] data, string? fontNameOverride) {
        Guard.NotNull(data, nameof(data));
        CacheBucket bucket = Cache.GetValue(data, static _ => new CacheBucket());
        string key = fontNameOverride ?? string.Empty;
        PdfTrueTypeFontProgram blueprint;
        lock (bucket.SyncRoot) {
            if (!bucket.TrueType.TryGetValue(key, out Lazy<PdfTrueTypeFontProgram>? lazy)) {
                lazy = new Lazy<PdfTrueTypeFontProgram>(
                    () => PdfTrueTypeFontProgram.Parse(data, fontNameOverride),
                    System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);
                bucket.TrueType[key] = lazy;
            }
            blueprint = lazy.Value;
        }
        return blueprint.ForkForDocument();
    }

    internal static PdfOpenTypeCffFontProgram GetOpenTypeCff(byte[] data, string? fontNameOverride) {
        Guard.NotNull(data, nameof(data));
        CacheBucket bucket = Cache.GetValue(data, static _ => new CacheBucket());
        string key = fontNameOverride ?? string.Empty;
        PdfOpenTypeCffFontProgram blueprint;
        lock (bucket.SyncRoot) {
            if (!bucket.OpenTypeCff.TryGetValue(key, out Lazy<PdfOpenTypeCffFontProgram>? lazy)) {
                lazy = new Lazy<PdfOpenTypeCffFontProgram>(
                    () => PdfOpenTypeCffFontProgram.Parse(data, fontNameOverride),
                    System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);
                bucket.OpenTypeCff[key] = lazy;
            }
            blueprint = lazy.Value;
        }
        return blueprint.ForkForDocument();
    }

    private sealed class CacheBucket {
        internal object SyncRoot { get; } = new();
        internal Dictionary<string, Lazy<PdfTrueTypeFontProgram>> TrueType { get; } =
            new(StringComparer.Ordinal);
        internal Dictionary<string, Lazy<PdfOpenTypeCffFontProgram>> OpenTypeCff { get; } =
            new(StringComparer.Ordinal);
    }
}

using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Keeps simple rendered family identities while retaining CSS quoting for literal delimiters and escapes.</summary>
    private static string InlineFallbackFamilyRequest(OfficeFontFallbackRun fallback) =>
        IsSingleFamilyName(fallback.FamilyName) ? fallback.FamilyName : fallback.FamilyRequest;

    // Every inline run asks this of the few family names in a document; parse each name once.
    private const int MaximumCachedFamilyNames = 1024;
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, bool> SingleFamilyNames =
        new(StringComparer.Ordinal);
    private static int _cachedFamilyNames;

    private static bool IsSingleFamilyName(string familyName) {
        if (SingleFamilyNames.TryGetValue(familyName, out bool single)) return single;
        var parsed = OfficeFontFamilyParser.Parse(familyName, allowSemicolonSeparators: true);
        single = parsed.Count == 1 && string.Equals(parsed[0], familyName, StringComparison.Ordinal);
        if (System.Threading.Volatile.Read(ref _cachedFamilyNames) < MaximumCachedFamilyNames
            && SingleFamilyNames.TryAdd(familyName, single)) System.Threading.Interlocked.Increment(ref _cachedFamilyNames);
        return single;
    }
}

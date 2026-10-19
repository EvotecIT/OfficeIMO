using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Keeps simple rendered family identities while retaining CSS quoting for literal delimiters and escapes.</summary>
    private static string InlineFallbackFamilyRequest(OfficeFontFallbackRun fallback) {
        var parsed = OfficeFontFamilyParser.Parse(fallback.FamilyName, allowSemicolonSeparators: true);
        return parsed.Count == 1 && string.Equals(parsed[0], fallback.FamilyName, StringComparison.Ordinal)
            ? fallback.FamilyName
            : fallback.FamilyRequest;
    }
}

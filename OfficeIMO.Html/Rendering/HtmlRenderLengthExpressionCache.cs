using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

/// <summary>Reuses short immutable expressions within one layout resolver, never their context-dependent used values.</summary>
internal sealed class HtmlRenderLengthExpressionCache {
    private const int MaximumEntries = 128;
    private const int MaximumSourceLength = 128;
    private readonly Dictionary<string, HtmlCssMathParseResult> _expressions = new(StringComparer.Ordinal);

    // Only the owning, synchronous style resolver uses this scratch context.
    // Every supplied field is overwritten before resolution and neither the
    // parser expression nor the numeric result retains a reference to it.
    internal HtmlCssLengthResolutionContext ResolutionContext { get; } = new();

    internal HtmlCssMathParseResult Parse(string value) {
        bool eligible = value.Length <= MaximumSourceLength;
        if (eligible && _expressions.TryGetValue(value, out HtmlCssMathParseResult? found)) return found;

        // The existing parser remains the sole grammar and limit owner. Once the
        // small operation-local cache is full, other values follow the usual path.
        HtmlCssMathParseResult parsed = HtmlCssMathParser.ParseLengthPercentage(value);
        if (eligible && parsed.IsParsed && _expressions.Count < MaximumEntries) {
            _expressions.Add(value, parsed);
        }
        return parsed;
    }
}

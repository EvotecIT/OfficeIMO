using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

/// <summary>Reuses short immutable expressions within one layout resolver, never their context-dependent used values.</summary>
internal sealed class HtmlRenderLengthExpressionCache {
    private const int MaximumEntries = 128;
    private const int MaximumSourceLength = 128;
    private readonly Dictionary<string, HtmlCssMathParseResult> _expressions = new(StringComparer.Ordinal);
    // Rejected entries, oldest first. A full cache gives one up for a value that parses.
    private readonly Queue<string> _rejected = new();

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
        // A rejected keyword such as a table cell's vertical-align is just as immutable as an
        // expression, and is asked for again for every cell that inherits it.
        if (!eligible) return parsed;
        if (_expressions.Count >= MaximumEntries) {
            // Rejected values must not keep a repeated valid expression out of a full cache.
            if (!parsed.IsParsed || _rejected.Count == 0) return parsed;
            _expressions.Remove(_rejected.Dequeue());
        }
        _expressions.Add(value, parsed);
        if (!parsed.IsParsed) _rejected.Enqueue(value);
        return parsed;
    }
}

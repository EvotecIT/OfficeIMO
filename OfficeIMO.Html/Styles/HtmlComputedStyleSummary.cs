namespace OfficeIMO.Html;

/// <summary>
/// Compact summary of computed CSS used by conversion profiles, reports, and tests.
/// </summary>
public sealed class HtmlComputedStyleSummary {
    internal HtmlComputedStyleSummary(
        int elementCount,
        int styledElementCount,
        int hiddenElementCount,
        IEnumerable<string> propertyNames,
        IEnumerable<string> fontFamilies,
        IEnumerable<string> colorValues) {
        ElementCount = elementCount;
        StyledElementCount = styledElementCount;
        HiddenElementCount = hiddenElementCount;
        PropertyNames = ToReadOnlyList(propertyNames);
        FontFamilies = ToReadOnlyList(fontFamilies);
        ColorValues = ToReadOnlyList(colorValues);
    }

    /// <summary>Number of elements seen by the computed-style engine.</summary>
    public int ElementCount { get; }

    /// <summary>Number of elements with at least one computed property.</summary>
    public int StyledElementCount { get; }

    /// <summary>Number of elements hidden by effective display or visibility values.</summary>
    public int HiddenElementCount { get; }

    /// <summary>Distinct computed property names.</summary>
    public IReadOnlyList<string> PropertyNames { get; }

    /// <summary>Distinct font-family values discovered in computed styles.</summary>
    public IReadOnlyList<string> FontFamilies { get; }

    /// <summary>Distinct color-like values discovered in computed styles.</summary>
    public IReadOnlyList<string> ColorValues { get; }

    private static IReadOnlyList<string> ToReadOnlyList(IEnumerable<string> values) {
        return values
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToList()
            .AsReadOnly();
    }
}

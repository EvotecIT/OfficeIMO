namespace OfficeIMO.Html;

/// <summary>Controls public managed computed-style inspection.</summary>
public sealed class HtmlComputedStyleOptions {
    /// <summary>CSS media context used to select applicable rules.</summary>
    public HtmlCssMediaContext MediaContext { get; set; } = HtmlCssMediaContext.Screen;
    /// <summary>Retain provider-neutral cascade traces for properties in <see cref="Css.HtmlCssPropertyCatalog"/>.</summary>
    /// <remarks>Disabled by default to avoid retaining candidate graphs for every element.</remarks>
    public bool IncludeCascadeTraces { get; set; }
}

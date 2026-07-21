namespace OfficeIMO.Html;

/// <summary>Declared support outcome for a semantic feature and conversion target.</summary>
public enum HtmlCapabilitySupportLevel {
    /// <summary>The target preserves the feature through its documented native or rendered contract.</summary>
    Supported,
    /// <summary>The target retains useful evidence but may simplify or flatten the feature.</summary>
    Approximated,
    /// <summary>The target does not currently represent the feature.</summary>
    Unsupported
}

namespace OfficeIMO.Html;

/// <summary>Preferred color scheme exposed to CSS media queries during static rendering.</summary>
public enum HtmlPreferredColorScheme {
    /// <summary>Evaluate <c>(prefers-color-scheme: light)</c> as active.</summary>
    Light,
    /// <summary>Evaluate <c>(prefers-color-scheme: dark)</c> as active.</summary>
    Dark
}

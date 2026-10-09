namespace OfficeIMO.Epub;

/// <summary>Reusable reflowable-book stylesheet profiles. Reader support and source CSS affect presentation.</summary>
public enum EpubTypographyProfile {
    /// <summary>Responsive images, tables, links and preformatted text; the existing manuscript default.</summary>
    Basic,
    /// <summary>Prose spacing, paragraph indentation and heading/page-break hints, using reader-selected fonts and colors.</summary>
    Prose,
    /// <summary>Technical prose, tables and wrapping code blocks with theme-inherited borders and relative sizing.</summary>
    Technical
}

/// <summary>Creates CSS for the EPUB authoring and manuscript-import owners, without adding resources or invoking a renderer.</summary>
public static class EpubTypography {
    private const string BasicCss = "body{line-height:1.5}img,svg{max-width:100%;height:auto}table{border-collapse:collapse;max-width:100%}th,td{padding:.25em}pre{white-space:pre-wrap;overflow-wrap:anywhere}a{overflow-wrap:anywhere}";
    private const string ProseCss = "table{overflow-wrap:anywhere}body{line-height:1.6}p{margin-block:.65em;text-align:start;orphans:2;widows:2}p+p{text-indent:1.2em;margin-block-start:0}h1,h2,h3,h4,h5,h6{line-height:1.25;break-after:avoid}blockquote{margin-inline:1.2em}figure{margin-inline:0}figcaption{font-size:.9em}";
    private const string TechnicalCss = "p{margin-block:.75em;text-align:start}h1,h2,h3,h4,h5,h6{line-height:1.25;break-after:avoid}table{overflow-wrap:anywhere}caption{text-align:start;font-weight:bold}th,td{text-align:start;vertical-align:top;border:.06em solid currentColor}pre{padding:.6em;border:.06em solid currentColor;tab-size:4}code,kbd,samp{font-family:monospace;font-size:.9em}figure{margin-inline:0}figcaption{font-size:.9em}";

    /// <summary>
    /// Returns a dependency-free CSS baseline. Profiles do not set body font family/size, text/background colors,
    /// fixed page widths, or important declarations. Prose/technical profiles use logical spacing for text direction.
    /// Apply before publisher CSS when publisher declarations should override this baseline.
    /// </summary>
    public static string CreateStylesheet(EpubTypographyProfile profile = EpubTypographyProfile.Basic) => profile switch {
        EpubTypographyProfile.Basic => BasicCss,
        EpubTypographyProfile.Prose => BasicCss + ProseCss,
        EpubTypographyProfile.Technical => BasicCss + TechnicalCss,
        _ => throw new ArgumentOutOfRangeException(nameof(profile))
    };
}

namespace OfficeIMO.Html;

/// <summary>Selects how a native Office adapter interprets HTML input.</summary>
public enum HtmlImportMode {
    /// <summary>Use a compatible OfficeIMO semantic envelope when present; otherwise import ordinary HTML.</summary>
    Auto = 0,
    /// <summary>Require the adapter's versioned OfficeIMO semantic envelope and restoration metadata.</summary>
    Semantic = 1,
    /// <summary>Interpret ordinary headings, sections, paragraphs, lists, tables, and images without proprietary metadata.</summary>
    Generic = 2
}

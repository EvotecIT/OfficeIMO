namespace OfficeIMO.Html.Dom;

/// <summary>Document layout compatibility mode determined during HTML parsing.</summary>
public enum HtmlDocumentMode {
    /// <summary>Standards mode.</summary>
    Standards,
    /// <summary>Limited quirks mode.</summary>
    LimitedQuirks,
    /// <summary>Quirks mode.</summary>
    Quirks
}

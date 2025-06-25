namespace OfficeIMO.Word;

/// <summary>
/// Lists the built-in page numbering presets that can be applied to
/// headers or footers when numbering pages.
/// </summary>
public enum WordPageNumberStyle {
    /// <summary>
    /// Simple numeric page numbers.
    /// </summary>
    PlainNumber,
    /// <summary>
    /// Page numbers with an accent bar above.
    /// </summary>
    AccentBar,
    /// <summary>
    /// "Page X of Y" format.
    /// </summary>
    PageNumberXofY,
    /// <summary>
    /// Numbers enclosed in square brackets.
    /// </summary>
    Brackets1,
    /// <summary>
    /// Due to way this page number style is built the location is always header (center) regardless of placement in document
    /// </summary>
    Brackets2,
    /// <summary>
    /// Page numbers separated by dotted leaders.
    /// </summary>
    Dots,
    /// <summary>
    /// Large italic page numbers.
    /// </summary>
    LargeItalics,
    /// <summary>
    /// Roman numerals (I, II, III...).
    /// </summary>
    Roman,
    /// <summary>
    /// Numbers decorated with tildes.
    /// </summary>
    Tildes,
    /// <summary>
    /// Due to way this page number style is built the location is always footer regardless of placement in document
    /// </summary>
    TwoBars,
    /// <summary>
    /// Page numbers separated from text by a top line.
    /// </summary>
    TopLine,
    /// <summary>
    /// Page numbers preceded by a tab stop.
    /// </summary>
    Tab,
    /// <summary>
    /// Page numbers separated from text by a thick line.
    /// </summary>
    ThickLine,
    //ThinLine,
    /// <summary>
    /// Numbers enclosed in a rounded rectangle.
    /// </summary>
    RoundedRectangle,
    /// <summary>
    /// Due to way this page number style is built the location is always header (center) regardless of placement in document
    /// </summary>
    Circle,
    /// <summary>
    /// Due to way this page number style is built the location is always header (right) regardless of placement in document
    /// </summary>
    VeryLarge,
    /// <summary>
    /// Due to way this page number style is built the location is always header (left) regardless of placement in document
    /// </summary>
    VerticalOutline1,
    /// <summary>
    /// Due to way this page number style is built the location is always header (right) regardless of placement in document
    /// </summary>
    VerticalOutline2
}

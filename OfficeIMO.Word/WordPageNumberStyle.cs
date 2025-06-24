namespace OfficeIMO.Word;

public enum WordPageNumberStyle {
    PlainNumber,
    AccentBar,
    PageNumberXofY,
    Brackets1,
    /// <summary>
    /// Due to way this page number style is built the location is always header (center) regardless of placement in document
    /// </summary>
    Brackets2,
    Dots,
    LargeItalics,
    Roman,
    Tildes,
    /// <summary>
    /// Due to way this page number style is built the location is always footer regardless of placement in document
    /// </summary>
    TwoBars,
    TopLine,
    Tab,
    ThickLine,
    //ThinLine,
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

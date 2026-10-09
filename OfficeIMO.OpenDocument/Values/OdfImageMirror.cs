namespace OfficeIMO.OpenDocument;

/// <summary>Native image mirroring stored in graphic styles.</summary>
/// <remarks>Combine vertical mirroring with at most one horizontal mode.
/// Draw projection supports unconditional horizontal mirroring; other modes remain explicit losses.</remarks>
[Flags]
public enum OdfImageMirror {
    /// <summary>Disables inherited mirroring.</summary>
    None = 0,
    /// <summary>Mirrors horizontally on every page.</summary>
    Horizontal = 1,
    /// <summary>Mirrors vertically.</summary>
    Vertical = 2,
    /// <summary>Mirrors horizontally on odd numbered pages.</summary>
    HorizontalOnOddPages = 4,
    /// <summary>Mirrors horizontally on even numbered pages.</summary>
    HorizontalOnEvenPages = 8
}

namespace OfficeIMO.Epub;

/// <summary>Physical orientation requested from the reading system.</summary>
public enum EpubPageOrientation {
    /// <summary>Allow the reading system to choose.</summary>
    Auto,
    /// <summary>Request a landscape display.</summary>
    Landscape,
    /// <summary>Request a portrait display.</summary>
    Portrait
}

/// <summary>When a reading system should synthesize a two-page spread.</summary>
public enum EpubPageSpread {
    /// <summary>Allow the reading system to choose.</summary>
    Auto,
    /// <summary>Display this page alone.</summary>
    None,
    /// <summary>Use spreads only in landscape orientation.</summary>
    Landscape,
    /// <summary>Use spreads in either orientation.</summary>
    Both
}

/// <summary>Placement of a page within a synthetic spread.</summary>
public enum EpubPageSide {
    /// <summary>Use page progression and reading-system defaults.</summary>
    Auto,
    /// <summary>Place the page on the left.</summary>
    Left,
    /// <summary>Place the page on the right.</summary>
    Right,
    /// <summary>Center a single page, overriding synthetic spread placement.</summary>
    Center
}

/// <summary>EPUB 3 XHTML page canvas and reading-system presentation requests.</summary>
public sealed class EpubFixedLayoutPage {
    /// <summary>Creates a canvas measured in CSS pixels. Dimensions must be positive.</summary>
    public EpubFixedLayoutPage(int width, int height) { Width = width; Height = height; }
    /// <summary>Canvas width in CSS pixels.</summary>
    public int Width { get; }
    /// <summary>Canvas height in CSS pixels.</summary>
    public int Height { get; }
    /// <summary>
    /// Complete set of positioned top-level body elements, at most 1024. Empty removes previously
    /// generated region rules. Unselected elements retain normal CSS behavior; DOM order is preserved.
    /// </summary>
    public IReadOnlyList<EpubFixedLayoutRegion> Regions { get; set; } = Array.Empty<EpubFixedLayoutRegion>();
    /// <summary>Requested orientation, independent of the canvas aspect ratio.</summary>
    public EpubPageOrientation Orientation { get; set; }
    /// <summary>Requested synthetic spread behavior.</summary>
    public EpubPageSpread Spread { get; set; }
    /// <summary>Requested position in a synthetic spread.</summary>
    public EpubPageSide Side { get; set; }
}

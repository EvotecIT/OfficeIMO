namespace OfficeIMO.Rtf;

/// <summary>
/// Broad category for an RTF destination control word.
/// </summary>
public enum RtfDestinationType {
    /// <summary>The destination is not registered by OfficeIMO.Rtf yet.</summary>
    Unknown,

    /// <summary>Document header or reader configuration destination.</summary>
    Header,

    /// <summary>Document information metadata.</summary>
    Metadata,

    /// <summary>Font table destination.</summary>
    FontTable,

    /// <summary>Color table destination.</summary>
    ColorTable,

    /// <summary>Stylesheet destination.</summary>
    StyleSheet,

    /// <summary>List definition destination.</summary>
    ListTable,

    /// <summary>Picture payload destination.</summary>
    Picture,

    /// <summary>Embedded object destination.</summary>
    Object,

    /// <summary>Drawing shape destination.</summary>
    Drawing,

    /// <summary>Field destination.</summary>
    Field,

    /// <summary>Header, footer, or related page furniture destination.</summary>
    HeaderFooter,

    /// <summary>Footnote destination.</summary>
    Footnote,

    /// <summary>Endnote destination.</summary>
    Endnote,

    /// <summary>Annotation destination.</summary>
    Annotation,

    /// <summary>Bookmark start or end marker destination.</summary>
    Bookmark,

    /// <summary>Destination whose text participates in the visible document body.</summary>
    BodyText
}

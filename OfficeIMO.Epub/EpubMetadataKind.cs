namespace OfficeIMO.Epub;

/// <summary>Classifies an entry from the OPF metadata element.</summary>
public enum EpubMetadataKind {
    /// <summary>Dublin Core metadata such as title, creator, language, or identifier.</summary>
    DublinCore,

    /// <summary>EPUB meta metadata, including EPUB 3 refinements and EPUB 2 name/content entries.</summary>
    Meta,

    /// <summary>EPUB 3 linked metadata.</summary>
    Link,

    /// <summary>Another metadata element retained for forward compatibility.</summary>
    Other
}

namespace OfficeIMO.Epub;

/// <summary>Identifies the package structure that supplied an EPUB navigation item.</summary>
public enum EpubNavigationSource {
    /// <summary>The source was not identified.</summary>
    Unknown = 0,

    /// <summary>An EPUB 3 XHTML navigation document.</summary>
    Epub3Navigation,

    /// <summary>An EPUB 2 NCX document.</summary>
    Ncx,

    /// <summary>An EPUB 2 OPF guide reference.</summary>
    Epub2Guide
}

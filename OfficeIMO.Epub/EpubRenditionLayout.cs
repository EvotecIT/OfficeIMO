namespace OfficeIMO.Epub;

/// <summary>EPUB rendition layout declared globally or for a spine item.</summary>
public enum EpubRenditionLayout {
    /// <summary>Content is dynamically paginated by the reading system.</summary>
    Reflowable,

    /// <summary>Content is pre-paginated and has fixed page geometry.</summary>
    PrePaginated
}

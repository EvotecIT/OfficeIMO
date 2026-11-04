namespace OfficeIMO.Publisher.Internal;

/// <summary>Read-only imported graphics, native bindings and original storage. Absent for source-free publications; current semantic edits never replace these records.</summary>
internal sealed class PublisherImportedResidue {
    internal PublisherImportedResidue(PublisherContents contents, PublisherEscherData artwork, PublisherNativeSource? nativeSource = null,
        PublisherPublicationState? baseline = null, PublisherQuillStyleCatalog? styles = null) {
        Contents = contents; Artwork = artwork; NativeSource = nativeSource; Baseline = baseline;
        Styles = styles ?? PublisherQuillStyleCatalog.Empty;
    }
    internal PublisherContents Contents { get; }
    internal PublisherEscherData Artwork { get; }
    internal PublisherNativeSource? NativeSource { get; }
    /// <summary>The import-time semantic inputs. Derived layout and later edits never replace this baseline.</summary>
    internal PublisherPublicationState? Baseline { get; }
    internal PublisherQuillStyleCatalog Styles { get; }
}

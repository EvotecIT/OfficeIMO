namespace OfficeIMO.Publisher.Internal;

/// <summary>Read-only imported graphics, native bindings and original storage. Absent for source-free publications; current semantic edits never replace these records.</summary>
internal sealed class PublisherImportedResidue {
    internal PublisherImportedResidue(PublisherContents contents, PublisherEscherData artwork, PublisherNativeSource? nativeSource = null) {
        Contents = contents; Artwork = artwork; NativeSource = nativeSource;
    }
    internal PublisherContents Contents { get; }
    internal PublisherEscherData Artwork { get; }
    internal PublisherNativeSource? NativeSource { get; }
}

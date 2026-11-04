using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

// Synthetic native-record fixtures enter the same semantic importer and
// projector as loaded publications, without claiming original source bytes.
internal static class PublisherProjectionFixture {
    internal static PublisherDocument Project(PublisherContents source, PublisherQuillText text,
        PublisherEscherData artwork, PublisherParseContext context) {
        PublisherPublicationState publication = PublisherModelImporter.Import(source, text, artwork, context);
        return new PublisherPageProjector(publication, new PublisherImportedResidue(source, artwork), context).Project();
    }
}

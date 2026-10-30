using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

// Current semantic pages, pictures, tables, stories and frame chains feed projection.
// Unsupported imported graphics and native provenance stay separate and inert.
internal sealed partial class PublisherModelState {
    private readonly PublisherImportedResidue? _imported;
    private readonly PublisherParseContext _decoded;
    private readonly PublisherReadReport _sourceReport;
    private readonly PublisherPublicationState _publication;

    internal PublisherModelState(PublisherContents contents, PublisherQuillText text, PublisherEscherData artwork,
        PublisherParseContext decoded, PublisherNativeSource nativeSource, PublisherReadReport? sourceReport = null,
        PublisherPublicationState? publication = null) {
        _publication = publication ?? PublisherModelImporter.Import(contents, text, artwork, decoded);
        _imported = new PublisherImportedResidue(contents, artwork, nativeSource, _publication, text.StyleCatalog); _decoded = decoded;
        _sourceReport = sourceReport ?? new PublisherReadReport(decoded.Diagnostics, contents.Shapes.Count, 0,
            text.Stories.Count, artwork.Images.Count, decoded.Records);
    }

    private PublisherModelState(PublisherPublicationState publication, PublisherParseContext context,
        PublisherReadReport sourceReport, PublisherImportedResidue? imported = null) {
        _publication = publication; _decoded = context; _sourceReport = sourceReport; _imported = imported;
    }

    internal static PublisherModelState Create(double width, double height, PublisherReadOptions options, CancellationToken token) =>
        new(PublisherPublicationState.Create(width, height), new PublisherParseContext(options, token),
            new PublisherReadReport(Array.Empty<OfficeConversionFidelityDiagnostic>(), 0, 0, 0, 0, 0));

    private int SourceObjectCount => _imported?.Contents.Shapes.Count ?? 0;
    internal PublisherQuillStyleCatalog SourceStyleCatalog => _imported?.Styles ?? PublisherQuillStyleCatalog.Empty;

    internal PublisherNativePreparation PrepareNativeBytes(long maximumBytes, CancellationToken token) =>
        PublisherNativeWriter.Prepare(_publication, _imported, _decoded.Options.Limits, maximumBytes, token);

    internal PublisherDocument Project(CancellationToken token) {
        token.ThrowIfCancellationRequested();
        PublisherDocument document = new PublisherPageProjector(_publication, _imported,
            _decoded.ForProjection(token, _publication.AdditionalImageBytes, _publication.AdditionalImageCount)).Project();
        token.ThrowIfCancellationRequested();
        document.AttachModel(this, _sourceReport);
        return document;
    }

    internal PublisherModelState ReplaceStory(uint storyId, IReadOnlyList<OfficeRichTextParagraph> paragraphs, CancellationToken token) {
        if (!_publication.Text.Stories.TryGetValue(storyId, out PublisherTextStory? original))
            throw new ArgumentOutOfRangeException(nameof(storyId), "The publication has no such story.");
        if (_publication.Text.IsTableStory(storyId))
            throw new NotSupportedException("Use ReplaceTableCellParagraphs for a table story with qualified cell ownership.");
        PublisherTextStory replacement = PublisherStoryEditor.Replace(original, paragraphs, _decoded.Options, token);
        if (_publication.Text.Stories.Values.Sum(story => PublisherStoryEditor.CharacterCount(story.Id == storyId ? replacement : story)) > _decoded.Options.Limits.MaxTextCharacters)
            throw new ArgumentException("Publication story text exceeds the configured character limit.", nameof(paragraphs));
        return WithPublication(_publication.WithText(_publication.Text.ReplaceStory(replacement)));
    }

    internal PublisherModelState SetBounds(PublisherTextFrame frame, double x, double y, double width, double height) {
        return WithPublication(_publication.WithText(_publication.Text.SetBounds(frame.Id, x, y, width, height)));
    }

    internal PublisherModelState SetFrameLayout(uint id, int columns, double spacing, OfficeTextPadding padding, OfficeTextVerticalAlignment alignment) =>
        WithPublication(_publication.WithText(_publication.Text.SetFrameLayout(id, columns, spacing, padding, alignment, _decoded.Options)));

    internal PublisherModelState AddLinkedTextFrame(uint previousFrameId, uint pageId, double x, double y, double width, double height,
        CancellationToken token, out uint frameId) {
        token.ThrowIfCancellationRequested();
        if ((long)SourceObjectCount + _publication.AuthoredObjectCount >= _decoded.Options.Limits.MaxItems)
            throw new InvalidDataException("Publisher current object limit exceeded.");
        return WithPublication(_publication.AddLinkedTextFrame(previousFrameId, pageId, x, y, width, height, out frameId));
    }

    private PublisherModelState WithPublication(PublisherPublicationState publication) =>
        new(publication, _decoded, _sourceReport, _imported);
}

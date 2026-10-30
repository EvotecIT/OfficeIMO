using OfficeIMO.Core.Internal;
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
        _imported = new PublisherImportedResidue(contents, artwork, nativeSource); _decoded = decoded;
        _publication = publication ?? PublisherModelImporter.Import(contents, text, artwork, decoded);
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

    internal PublisherModelState AddLinkedTextFrame(uint previousFrameId, uint pageId, double x, double y, double width, double height,
        CancellationToken token, out uint frameId) {
        token.ThrowIfCancellationRequested();
        if ((long)SourceObjectCount + _publication.AuthoredObjectCount >= _decoded.Options.Limits.MaxItems)
            throw new InvalidDataException("Publisher current object limit exceeded.");
        return WithPublication(_publication.AddLinkedTextFrame(previousFrameId, pageId, x, y, width, height, out frameId));
    }

    internal PublisherModelState SetPictureBounds(uint pictureId, double x, double y, double width, double height) {
        PublisherPicture picture = _publication.EditablePicture(pictureId);
        return WithPublication(_publication.SetBounds(picture, x, y, width, height));
    }

    internal PublisherModelState ReplacePictureImage(uint pictureId, byte[] bytes, CancellationToken token) {
        PublisherPicture picture = _publication.EditablePicture(pictureId);
        if (bytes.Length == 0 || bytes.Length > _decoded.Options.MaximumImageBytes)
            throw new ArgumentException("Replacement image bytes must fit the configured per-image limit.", nameof(bytes));
        byte[] owned = (byte[])bytes.Clone();
        if (!OfficeImageReader.TryIdentifyByContent(owned, null, token, out OfficeImageInfo info)
            || info.Format is not (OfficeImageFormat.Png or OfficeImageFormat.Jpeg))
            throw new ArgumentException("Replacement pictures require PNG or JPEG content.", nameof(bytes));
        if (info.Width <= 0 || info.Height <= 0 || (long)info.Width * info.Height > _decoded.Options.MaximumRasterPixels)
            throw new ArgumentException("Replacement image dimensions exceed the configured pixel limit.", nameof(bytes));
        var options = new OfficeRasterDecodeOptions {
            MaximumEncodedBytes = Math.Min(_decoded.Options.MaximumImageBytes, 128 * 1024 * 1024),
            MaximumDecodedPixels = Math.Min(_decoded.Options.MaximumRasterPixels, 50_000_000),
            MaximumInspectionWorkPixels = Math.Min(_decoded.Options.MaximumImageProcessingPixels, 50_000_000),
            CancellationToken = token
        };
        if (!OfficeRasterImageDecoder.TryDecode(owned, options, out OfficeRasterImage? raster, out _) || raster == null)
            throw new ArgumentException("Replacement picture content is invalid or exceeds the configured decoding bounds.", nameof(bytes));
        token.ThrowIfCancellationRequested();
        PublisherPublicationState candidate = _publication.ReplaceImage(picture, owned, info.MimeType);
        if (candidate.AdditionalImageBytes > _decoded.Options.MaximumTotalImageBytes - _decoded.ImageBytes)
            throw new ArgumentException("Replacement image exceeds the configured publication image byte limit.", nameof(bytes));
        return WithPublication(candidate);
    }

    private PublisherModelState WithPublication(PublisherPublicationState publication) =>
        new(publication, _decoded, _sourceReport, _imported);
}

// Owned source storage is private to the model. Retention is not a native-save
// or unknown-record preservation claim; those operations need separate proof.
internal sealed class PublisherNativeSource {
    private readonly byte[] _bytes;
    private readonly OfficeCompoundFile _compound;
    internal PublisherNativeSource(byte[] bytes, OfficeCompoundFile compound) { _bytes = (byte[])bytes.Clone(); _compound = compound; }
    internal int ByteLength => _bytes.Length;
    internal int StreamCount => _compound.Streams.Count;
}

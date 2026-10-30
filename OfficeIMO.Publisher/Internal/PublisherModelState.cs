using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

// The document's current decoded model is the only input to projection. Native
// bytes are retained separately as immutable provenance, never patched by edits.
internal sealed class PublisherModelState {
    private readonly PublisherContents _contents;
    private readonly PublisherQuillText _text;
    private readonly PublisherEscherData _artwork;
    private readonly PublisherParseContext _decoded;
    private readonly PublisherNativeSource _nativeSource;
    private readonly PublisherReadReport _sourceReport;

    internal PublisherModelState(PublisherContents contents, PublisherQuillText text, PublisherEscherData artwork,
        PublisherParseContext decoded, PublisherNativeSource nativeSource, PublisherReadReport? sourceReport = null) {
        _contents = contents; _text = text; _artwork = artwork; _decoded = decoded; _nativeSource = nativeSource;
        _sourceReport = sourceReport ?? new PublisherReadReport(decoded.Diagnostics, contents.Shapes.Count, 0,
            text.Stories.Count, artwork.Images.Count, decoded.Records);
    }

    internal PublisherDocument Project(CancellationToken token) {
        token.ThrowIfCancellationRequested();
        PublisherDocument document = new PublisherPageProjector(_contents, _text, _artwork, _decoded.ForProjection(token)).Project();
        token.ThrowIfCancellationRequested();
        document.AttachModel(this, _sourceReport);
        return document;
    }

    internal PublisherModelState ReplaceStory(uint storyId, IReadOnlyList<OfficeRichTextParagraph> paragraphs, CancellationToken token) {
        if (!_text.Stories.TryGetValue(storyId, out PublisherTextStory? original))
            throw new ArgumentOutOfRangeException(nameof(storyId), "The publication has no such story.");
        if (_text.CellParagraphs.ContainsKey(storyId) || _contents.Shapes.Values.Any(shape => shape.Chunk.Kind == 0x10 && shape.Value(0x27) == storyId))
            throw new NotSupportedException("Editing table stories requires a qualified cell-to-story mapping and is not supported.");
        PublisherTextStory replacement = PublisherStoryEditor.Replace(original, paragraphs, _decoded.Options, token);
        if (_text.Stories.Values.Sum(story => PublisherStoryEditor.CharacterCount(story.Id == storyId ? replacement : story)) > _decoded.Options.Limits.MaxTextCharacters)
            throw new ArgumentException("Publication story text exceeds the configured character limit.", nameof(paragraphs));
        var text = new PublisherQuillText();
        foreach (var entry in _text.Stories) text.Stories.Add(entry.Key, entry.Key == storyId ? replacement : entry.Value);
        foreach (var entry in _text.CellParagraphs) text.CellParagraphs.Add(entry.Key, entry.Value);
        return new PublisherModelState(_contents, text, _artwork, _decoded, _nativeSource, _sourceReport);
    }

    internal PublisherModelState SetBounds(PublisherTextFrame frame, double x, double y, double width, double height) {
        PublisherEscherShape? shape = _artwork.Shapes.SingleOrDefault(item => item.Id == frame.Id);
        if (shape == null || !shape.Bounds.HasValue || shape.Hidden || shape.IsGroup || shape.HasPublicationGroupParent
            || shape.GroupTransform != OfficeTransform.Identity || shape.LayoutTransform != OfficeTransform.Identity
            || shape.Transform.RotationDegrees.GetValueOrDefault() != 0 || shape.Transform.FlipHorizontal || shape.Transform.FlipVertical)
            throw new NotSupportedException("Frame bounds can only be edited for visible, resolved, untransformed top-level frames.");
        PublisherSourcePage page = _contents.Pages.Single(item => item.Id == frame.PageId);
        double originX = _contents.Width / 2, originY = _contents.Height / 2;
        if (page.IsMaster) {
            if (page.NativeLayout.HasValue && page.NativeLayout.Value is not (0 or 1 or 3))
                throw new NotSupportedException("This master page's coordinate basis is not qualified for frame edits.");
            originX = page.NativeLayout switch { 0 => 0, 1 => _contents.Width, _ => originX };
        }
        double x1 = (x - originX) * 12700, y1 = (y - originY) * 12700;
        double x2 = (x + width - originX) * 12700, y2 = (y + height - originY) * 12700;
        if (new[] { x1, y1, x2, y2 }.Any(value => value < int.MinValue || value > int.MaxValue))
            throw new ArgumentOutOfRangeException(nameof(width), "Frame bounds exceed native signed coordinate limits.");
        var replacement = new PublisherEscherShape(shape.Id, shape.Type, shape.Flags, shape.Properties, shape.HasParentDrawingSpace) {
            Bounds = new PublisherNativeRectangle(x1, y1, x2, y2), GroupCoordinates = shape.GroupCoordinates,
            GroupTransform = shape.GroupTransform, LayoutTransform = shape.LayoutTransform,
            HasPublicationGroupParent = shape.HasPublicationGroupParent, HiddenByGroup = shape.HiddenByGroup
        };
        var artwork = new PublisherEscherData();
        foreach (PublisherEscherShape item in _artwork.Shapes) artwork.Shapes.Add(item.Id == frame.Id ? replacement : item);
        foreach (var entry in _artwork.Images) artwork.Images.Add(entry.Key, entry.Value);
        return new PublisherModelState(_contents, _text, artwork, _decoded, _nativeSource, _sourceReport);
    }
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

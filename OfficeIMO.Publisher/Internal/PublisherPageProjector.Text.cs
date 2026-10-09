using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private OfficeRasterCanvas? _textMetrics;

    private void ProjectText(IReadOnlyList<OfficeRichTextParagraph> paragraphs, OfficeDrawing drawing,
        PublisherEscherShape source, double x, double y, double width, double height,
        OfficeTextVerticalAlignment vertical = OfficeTextVerticalAlignment.Top) {
        OfficeTextPadding padding = TextPadding(source);
        string location = PublisherEscherReader.ShapeLocation(source.Id);
        if (padding.Horizontal >= width || padding.Vertical >= height) {
            _context.Add("PUB_TEXT_FRAME_HAS_NO_CONTENT_AREA", "Native text insets consume its frame. The text remains available in TextStories but has no printable content area.",
                OfficeConversionLossKind.Omission, location);
            return;
        }
        drawing.AddRichTextParagraphs(paragraphs, x, y, width, height, verticalAlignment: vertical, padding: padding);
        var frame = (OfficeDrawingRichText)drawing.Elements[drawing.Elements.Count - 1];
        // The shared managed metrics match the default SVG layout. PDF or
        // application fonts can produce different wrapping, so retain the
        // layout approximation report even when this measurement fits.
        _textMetrics ??= new OfficeRasterCanvas(new OfficeRasterImage(1, 1), null, null,
            cancellationToken: _context.Token);
        OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.CreateWithRasterMetrics(frame,
            width - padding.Horizontal, height - padding.Vertical, _textMetrics);
        if (layout.Clipped)
            _context.Add("PUB_TEXT_FRAME_OVERFLOW", "Shared managed text measurement clips part of this story to its frame. TextStories retains the complete text; output font metrics may change the clipped range.",
                OfficeConversionLossKind.Omission, location);
    }
}

using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    /// <summary>Paints a current authored frame without native shape descriptors or synthetic OfficeArt records.</summary>
    private void ProjectAuthoredTextFrame(PublisherFrameState definition, OfficeDrawing page) {
        var local = new OfficeDrawing(definition.Width, definition.Height);
        if (_textFrames.TryGetValue(definition.Id, out PreparedTextFrame? frame))
            foreach (PreparedTextRegion region in frame.Regions) {
                local.AddRichTextParagraphs(region.Paragraphs,
                    region.X, region.Y, region.Width, region.Height, verticalAlignment: definition.VerticalAlignment);
                ((OfficeDrawingRichText)local.Elements[local.Elements.Count - 1]).NativePublisherFrame = region.NativePublisherFrame;
            }
        if (definition.X < 0 || definition.Y < 0 || definition.X + definition.Width > page.Width || definition.Y + definition.Height > page.Height)
            _context.Add("PUB_PAGE_BLEED_CLIPPED", "Artwork outside the publication page remains in the scene but is clipped to the printable page in exports.",
                OfficeConversionLossKind.Approximation, "publication/frame/" + definition.Id);
        if (local.Elements.Count != 0) {
            TagSourceElements(local, new[] { "publisher-object-" + definition.Id });
            _context.AccountProjection(local);
            page.AddDrawingForClippedRendering(local, definition.X, definition.Y, null);
            _context.Add("PUB_TEXT_LAYOUT_APPROXIMATED", "Publication text uses shared drawing paragraph layout. Font availability, line breaking, fitting and overflow can differ from Publisher.",
                OfficeConversionLossKind.Approximation, "publication/story/" + definition.StoryId);
        }
        _projected.Add(definition.Id);
    }
}

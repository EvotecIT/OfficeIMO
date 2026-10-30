using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private readonly Dictionary<uint, PreparedTextFrame> _textFrames = new();

    private void PrepareTextFrames() {
        foreach (var chain in _publication.Text.Chains) {
            _context.Token.ThrowIfCancellationRequested();
            if (!_publication.Text.Stories.TryGetValue(chain.Key, out PublisherTextStory? story)) continue;
            var flow = new OfficeRichTextFlow(story.Paragraphs, _context.AccountTextLayout);
            bool unavailable = false;
            for (int frameIndex = 0; frameIndex < chain.Value.Count; frameIndex++) {
                PublisherFrameState frame = _publication.Text.Frames[chain.Value[frameIndex]];
                uint? previous = frameIndex == 0 ? null : chain.Value[frameIndex - 1];
                uint? next = frameIndex + 1 == chain.Value.Count ? null : chain.Value[frameIndex + 1];
                if (!frame.IsAuthored) _context.Record();
                if (!frame.Available) {
                    unavailable = true;
                    _context.Add("PUB_TEXT_CHAIN_FRAME_UNAVAILABLE", "A story frame has no printable resolved drawing. Continuation after this frame remains unplaced; the complete story is retained.",
                        OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(frame.Id));
                    continue;
                }
                int columns = frame.ColumnCount;
                OfficeTextPadding padding = frame.Padding;
                double width = (frame.Width - padding.Horizontal - (columns - 1) * frame.ColumnSpacing) / columns;
                double height = frame.Height - padding.Vertical;
                var content = new List<PreparedTextRegion>();
                IReadOnlyList<OfficeTextFlowRegion> exclusions = TextExclusions(frame);
                int? start = unavailable ? null : flow.HasRemaining ? Math.Min(flow.CharacterPosition, story.Text.Length) : story.Text.Length;
                if (!unavailable && width > 0 && height > 0) {
                    _textMetrics ??= new OfficeRasterCanvas(new OfficeRasterImage(1, 1), null, null, cancellationToken: _context.Token);
                    for (int column = 0; column < columns && flow.HasRemaining; column++) {
                        if (!frame.IsAuthored) _context.Record();
                        var area = new OfficeTextFlowRegion(padding.Left + column * (width + frame.ColumnSpacing), padding.Top, width, height);
                        IReadOnlyList<OfficeTextFlowRegion> regions = exclusions.Count == 0 ? new[] { area } :
                            OfficeTextFlowRegions.Exclude(area, exclusions, _context.Record, _context.Token);
                        foreach (OfficeTextFlowRegion region in regions) {
                            if (!flow.HasRemaining) break;
                            IReadOnlyList<OfficeRichTextParagraph> paragraphs = flow.Take(region.Width, region.Height,
                                _textMetrics.MeasureText, _context.Token);
                            if (paragraphs.Count != 0) content.Add(new PreparedTextRegion(paragraphs, region.X, region.Y, region.Width, region.Height));
                        }
                    }
                } else if (!unavailable) _context.Add("PUB_TEXT_FRAME_HAS_NO_CONTENT_AREA",
                    "Native text insets and column gaps consume the frame. Its text continues into the next linked frame when available.",
                    OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(frame.Id));
                int? end = unavailable ? null : flow.HasRemaining ? Math.Min(flow.CharacterPosition, story.Text.Length) : story.Text.Length;
                bool overflow = !next.HasValue && flow.HasRemaining;
                if (overflow) _context.Add("PUB_TEXT_FRAME_OVERFLOW",
                    "Shared managed measurement leaves story content beyond the final frame. TextStories retains the complete text; output font metrics may change the assigned range.",
                    OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(frame.Id));
                var model = new PublisherTextFrame(frame.Id, story.Id, frame.PageId!.Value, previous, next, (uint)frameIndex,
                    frame.X, frame.Y, frame.Width, frame.Height, columns, frame.ColumnSpacing, start, end - start, overflow,
                    frame.WrapObjectIds, frame.PageTransform);
                _textFrames.Add(frame.Id, new PreparedTextFrame(model, content));
                if (!unavailable) _usedStories.Add(story.Id);
            }
            if (chain.Value.Count > 1) {
                bool imported = chain.Value.Any(id => !_publication.Text.Frames[id].IsAuthored);
                _context.Add("PUB_LINKED_TEXT_LAYOUT_APPROXIMATED",
                    imported ? "Text follows the recovered native story-frame chain using shared managed measurement. Source paragraph styles survive continuation; native font metrics, column balancing and frame break positions can differ."
                        : "Text follows the current story-frame chain using shared managed measurement. Paragraph styles survive continuation; font metrics and frame break positions can differ from Publisher.",
                    OfficeConversionLossKind.Approximation, (imported ? "Quill/story/" : "publication/story/") + story.Id);
            }
        }
    }

    private sealed record PreparedTextFrame(PublisherTextFrame Model, IReadOnlyList<PreparedTextRegion> Regions);
    private sealed record PreparedTextRegion(IReadOnlyList<OfficeRichTextParagraph> Paragraphs, double X, double Y, double Width, double Height);
}

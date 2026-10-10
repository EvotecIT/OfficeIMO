using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private readonly Dictionary<uint, PreparedTextFrame> _textFrames = new();

    private void PrepareTextFrames() {
        var artwork = _escher.Shapes.ToDictionary(shape => shape.Id);
        var stories = _source.Shapes.Values.Where(shape => shape.Chunk.Kind != 0x10 && shape.Value(0x27).HasValue)
            .GroupBy(shape => shape.Value(0x27)!.Value);
        foreach (var references in stories) {
            _context.Record();
            if (!_text.Stories.TryGetValue(references.Key, out PublisherTextStory? story)) continue;
            PublisherSourceShape[] frames = OrderFrames(references.ToArray());
            var flow = new OfficeRichTextFlow(story.Paragraphs, _context.AccountTextLayout);
            bool unavailable = false;
            foreach (PublisherSourceShape frame in frames) {
                _context.Record();
                uint? owner = Owner(frame.Chunk.Id, 0);
                if (!owner.HasValue || !artwork.TryGetValue(frame.Chunk.Id, out PublisherEscherShape? shape)
                    || !shape.Bounds.HasValue || shape.IsGroup || shape.Hidden) {
                    // Subsequent frames cannot safely be assigned the missing frame's text.
                    unavailable = true;
                    _context.Add("PUB_TEXT_CHAIN_FRAME_UNAVAILABLE", "A story frame has no printable resolved drawing. Continuation after this frame remains unplaced; the complete story is retained.",
                        OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(frame.Chunk.Id));
                    continue;
                }
                FrameRectangle bounds = FrameBounds(shape, _source.Width, _source.Height);
                uint columns = shape.Property(0x8C) ?? 1;
                if (columns == 0 || columns > 256 || columns > _context.Options.Limits.MaxItems)
                    throw new InvalidDataException("Publisher text column count is outside the supported resource bounds.");
                double gap = (shape.Property(0x8D) ?? 0) / 12700D;
                OfficeTextPadding padding = TextPadding(shape);
                double width = (bounds.Width - padding.Horizontal - (columns - 1) * gap) / columns;
                double height = bounds.Height - padding.Vertical;
                var content = new List<PreparedTextRegion>();
                IReadOnlyList<OfficeTextFlowRegion> exclusions = TextExclusions(frame, shape, bounds, artwork);
                int? start = unavailable ? null : flow.HasRemaining ? Math.Min(flow.CharacterPosition, story.Text.Length) : story.Text.Length;
                if (!unavailable && width > 0 && height > 0) {
                    _textMetrics ??= new OfficeRasterCanvas(new OfficeRasterImage(1, 1), null, null, cancellationToken: _context.Token);
                    for (int column = 0; column < columns && flow.HasRemaining; column++) {
                        _context.Record();
                        var area = new OfficeTextFlowRegion(padding.Left + column * (width + gap), padding.Top, width, height);
                        IReadOnlyList<OfficeTextFlowRegion> regions = exclusions.Count == 0 ? new[] { area } :
                            OfficeTextFlowRegions.Exclude(area, exclusions, _context.Record, _context.Token);
                        foreach (OfficeTextFlowRegion region in regions) {
                            if (!flow.HasRemaining) break;
                            IReadOnlyList<OfficeRichTextParagraph> paragraphs = flow.Take(region.Width, region.Height,
                                _textMetrics.MeasureText, _context.Token);
                            if (paragraphs.Count != 0) content.Add(new PreparedTextRegion(paragraphs,
                                region.X, region.Y, region.Width, region.Height));
                        }
                    }
                } else if (!unavailable) _context.Add("PUB_TEXT_FRAME_HAS_NO_CONTENT_AREA",
                    "Native text insets and column gaps consume the frame. Its text continues into the next linked frame when available.",
                    OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(frame.Chunk.Id));
                int? end = unavailable ? null : flow.HasRemaining ? Math.Min(flow.CharacterPosition, story.Text.Length) : story.Text.Length;
                bool overflow = !NextFrame(frame).HasValue && flow.HasRemaining;
                if (overflow) _context.Add("PUB_TEXT_FRAME_OVERFLOW",
                    "Shared managed measurement leaves story content beyond the final frame. TextStories retains the complete text; output font metrics can change the assigned range.",
                    OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(frame.Chunk.Id));
                var model = new PublisherTextFrame(frame.Chunk.Id, story.Id, owner.Value, PreviousFrame(frame), NextFrame(frame),
                    frame.Value(0x28) ?? 0, bounds.X, bounds.Y, bounds.Width, bounds.Height,
                    (int)columns, gap, start, end - start, overflow, frame.WrapObjects,
                    OfficeTransform.Translate(bounds.X, bounds.Y).Then(PageTransform(shape, bounds)));
                _textFrames.Add(frame.Chunk.Id, new PreparedTextFrame(model, content));
                if (!unavailable) _usedStories.Add(story.Id);
            }
            if (frames.Length > 1) _context.Add("PUB_LINKED_TEXT_LAYOUT_APPROXIMATED",
                "Text follows native frame links using shared managed measurement. Source paragraph styles survive continuation; native font metrics, column balancing and frame break positions can differ.",
                OfficeConversionLossKind.Approximation, "Quill/story/" + story.Id);
        }
    }

    private PublisherSourceShape[] OrderFrames(PublisherSourceShape[] frames) {
        var byId = frames.ToDictionary(frame => frame.Chunk.Id);
        foreach (PublisherSourceShape frame in frames) {
            uint? previous = PreviousFrame(frame), next = NextFrame(frame);
            if (previous.HasValue && (!byId.TryGetValue(previous.Value, out PublisherSourceShape? before)
                || NextFrame(before) != frame.Chunk.Id)) throw new InvalidDataException("Publisher previous text-frame link is unresolved or not reciprocal.");
            if (next.HasValue && (!byId.TryGetValue(next.Value, out PublisherSourceShape? after)
                || PreviousFrame(after) != frame.Chunk.Id)) throw new InvalidDataException("Publisher next text-frame link is unresolved or not reciprocal.");
        }
        PublisherSourceShape[] roots = frames.Where(frame => !PreviousFrame(frame).HasValue).ToArray();
        if (roots.Length == 0) throw new InvalidDataException("Cyclic Publisher text-frame chain.");
        if (roots.Length != 1) {
            _context.Add("PUB_TEXT_CHAIN_UNRESOLVED", "A source story has multiple unlinked roots. Frame order is ambiguous, so its text is retained without guessed placement.",
                OfficeConversionLossKind.Omission, "Quill/story/" + frames[0].Value(0x27));
            return Array.Empty<PublisherSourceShape>();
        }
        var ordered = new List<PublisherSourceShape>();
        var visited = new HashSet<uint>();
        PublisherSourceShape? current = roots[0];
        while (current != null) {
            _context.Record();
            if (!visited.Add(current.Chunk.Id)) throw new InvalidDataException("Cyclic Publisher text-frame chain.");
            uint ordinal = current.Value(0x28) ?? 0;
            if (ordinal != ordered.Count) throw new InvalidDataException("Publisher text-frame ordinal disagrees with its native links.");
            ordered.Add(current);
            uint? next = NextFrame(current);
            current = next.HasValue ? byId[next.Value] : null;
        }
        if (ordered.Count != frames.Length) throw new InvalidDataException("Publisher text-frame story contains a disconnected cycle.");
        return ordered.ToArray();
    }

    private static uint? PreviousFrame(PublisherSourceShape frame) => FrameLink(frame, 0x36);
    private static uint? NextFrame(PublisherSourceShape frame) => FrameLink(frame, 0x37);
    private static uint? FrameLink(PublisherSourceShape frame, byte id) {
        PublisherBlock? block = PublisherContentsReader.Field(frame.Fields, id);
        if (!block.HasValue) return null;
        if (block.Value.Type is not (0x68 or 0x70)) throw new InvalidDataException("Invalid Publisher text-frame link encoding.");
        return block.Value.Value == 0 ? null : block.Value.Value;
    }

    private static FrameRectangle FrameBounds(PublisherEscherShape shape, double pageWidth, double pageHeight) {
        PublisherNativeRectangle native = shape.Bounds!.Value.Unrotated(shape.Transform.RotationDegrees.GetValueOrDefault());
        double x = Math.Min(native.X1, native.X2) / 12700 + pageWidth / 2;
        double y = Math.Min(native.Y1, native.Y2) / 12700 + pageHeight / 2;
        double width = Math.Abs(native.X2 - native.X1) / 12700, height = Math.Abs(native.Y2 - native.Y1) / 12700;
        return new FrameRectangle(x, y, width, height);
    }

    private readonly record struct FrameRectangle(double X, double Y, double Width, double Height);
    private sealed record PreparedTextFrame(PublisherTextFrame Model, IReadOnlyList<PreparedTextRegion> Regions);
    private sealed record PreparedTextRegion(IReadOnlyList<OfficeRichTextParagraph> Paragraphs, double X, double Y, double Width, double Height);
}

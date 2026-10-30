using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherTextState {
    // Projection supplies its known prefix state; snapshots must not each walk
    // the chain again. A mutation checks the prefix once before replacement.
    internal string? FrameLayoutEditLimitation(uint frameId, bool hasResolvedRange) {
        if (!Frames.TryGetValue(frameId, out PublisherFrameState? frame)) return "The publication has no such text frame.";
        if (IsTableStory(frame.StoryId)) return "Frame layout editing does not support table stories or their ordinary-frame aliases.";
        if (!frame.Available || !frame.CanEditBounds)
            return "Frame layout requires visible, resolved, untransformed top-level placement on a qualified page.";
        if (frame.LayoutEditLimitation != null) return frame.LayoutEditLimitation;
        if (!Stories.ContainsKey(frame.StoryId)) return "The frame has no complete story.";
        return hasResolvedRange ? null : "An unresolved chain or unavailable earlier frame leaves this story range unresolved.";
    }

    internal PublisherTextState SetFrameLayout(uint frameId, int columns, double spacing, OfficeTextPadding padding,
        OfficeTextVerticalAlignment alignment, PublisherReadOptions options) {
        if (!Frames.TryGetValue(frameId, out PublisherFrameState? frame))
            throw new ArgumentOutOfRangeException(nameof(frameId), "The publication has no such text frame.");
        string? limitation = FrameLayoutEditLimitation(frameId, HasResolvedFrameRange(frame));
        if (limitation != null) throw new NotSupportedException(limitation);
        if (columns > options.Limits.MaxItems) throw new InvalidDataException("Publisher text column count exceeds the configured item limit.");
        PublisherFrameState replacement = frame.WithLayout(columns, spacing, padding, alignment);
        // Existing collapsed native insets remain readable and can be corrected.
        // Validate the requested area, never use the old area as an edit gate.
        if (!Finite(padding.Horizontal) || !Finite(padding.Vertical) || !Finite((columns - 1) * spacing)
            || !Finite(replacement.ColumnWidth) || !Finite(replacement.ContentHeight)
            || replacement.ColumnWidth <= 0 || replacement.ContentHeight <= 0)
            throw new ArgumentException("Text-frame padding and column gaps must leave a positive finite content area.", nameof(padding));
        var frames = Frames.ToDictionary(item => item.Key, item => item.Value); frames[frameId] = replacement;
        return new PublisherTextState(this, frames: frames);
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private bool HasResolvedFrameRange(PublisherFrameState frame) {
        if (!Chains.TryGetValue(frame.StoryId, out IReadOnlyList<uint>? chain)) return false;
        foreach (uint id in chain) {
            if (!Frames[id].Available) return false;
            if (id == frame.Id) return true;
        }
        return false;
    }
}

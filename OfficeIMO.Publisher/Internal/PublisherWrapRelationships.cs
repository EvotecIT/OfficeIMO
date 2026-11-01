using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Derives ordinary square-wrap frame relationships from current page geometry and stacking.</summary>
internal static class PublisherWrapRelationships {
    internal static PublisherNativeWrapMode ReadMode(PublisherSourceShape? source, PublisherEscherShape artwork,
        bool picture, bool negativeMargins, bool tight) {
        if (source == null) return PublisherNativeWrapMode.Unknown;
        PublisherBlock[] mode = source.Fields.Where(field => field.Id == 4).ToArray();
        PublisherBlock[] distance = source.Fields.Where(field => field.Id == 0x12).ToArray();
        if (mode.Length > 1 || mode.Any(field => field.Type != 0x10)
            || distance.Length > 1 || distance.Any(field => field.Type != 0)) return PublisherNativeWrapMode.Unknown;
        if (mode.Length == 1 && mode[0].Value == 0x100)
            return PublisherNativeWrapMode.None;
        // Native side controls use field05; only the omitted both-side default
        // and rectangular, non-outline, nonnegative-margin obstacles are qualified.
        // An omitted mode is native square wrapping (also with automatic distances);
        // ordinary rectangle/text-box through wrapping has the same solid contour.
        if (!negativeMargins && !tight && source.Fields.All(field => field.Id != 5)
            && !OfficeArtCustomPathProjector.HasPath(artwork.Properties)
            && (picture || artwork.Type is 1 or 202)
            && (mode.Length == 0 || !picture && mode[0].Value == 0x103))
            return PublisherNativeWrapMode.RectangularBoth;
        return PublisherNativeWrapMode.Unknown;
    }

    internal static bool GeometryChanged(PublisherPublicationState before, PublisherPageState a,
        PublisherPublicationState after, PublisherPageState b) {
        if (!a.ObjectIds.SequenceEqual(b.ObjectIds)) return true;
        return a.ObjectIds.Any(id => Placement(before, id) != Placement(after, id));
    }

    internal static bool Matches(PublisherPublicationState state, IReadOnlyDictionary<uint, IReadOnlyList<uint>> wraps) =>
        wraps.All(pair => state.Text.Frames[pair.Key].WrapObjectIds.SequenceEqual(pair.Value));

    internal static bool TryCompute(PublisherPublicationState state, PublisherPageState page, PublisherParseContext context,
        out IReadOnlyDictionary<uint, IReadOnlyList<uint>> relationships) {
        var result = new Dictionary<uint, IReadOnlyList<uint>>(); relationships = result;
        var recipients = page.ObjectIds.Where(state.Text.Frames.ContainsKey).ToArray();
        if (recipients.Length == 0) return true;
        if (page.ObjectIds.All(id => state.Text.WrapInfo.TryGetValue(id, out var info)
            ? info.NativeMode == PublisherNativeWrapMode.None : Authored(state, id))) {
            foreach (uint id in recipients) { context.Record(); result.Add(id, Array.Empty<uint>()); }
            return true;
        }
        var positions = page.ObjectIds.Select((id, index) => (id, index)).ToDictionary(pair => pair.id, pair => pair.index);
        foreach (uint frameId in recipients) {
            context.Record(); PublisherFrameState frame = state.Text.Frames[frameId];
            if (page.ObjectIds.Count == 1) { result.Add(frameId, Array.Empty<uint>()); continue; }
            if (!frame.CanEditBounds || frame.Hidden || !frame.HasAnchor || !frame.TransformAvailable
                || state.Text.IsTableStory(frame.StoryId)) return false;
            Bounds? target = Placement(state, frameId);
            if (!target.HasValue) return false;
            var ids = new List<uint>();
            foreach (uint id in page.ObjectIds.Skip(positions[frameId] + 1)) {
                context.Record();
                PublisherNativeWrapMode mode = state.Text.WrapInfo.TryGetValue(id, out PublisherWrapInfo? wrap)
                    ? wrap.NativeMode : Authored(state, id) ? PublisherNativeWrapMode.None : PublisherNativeWrapMode.Unknown;
                if (mode == PublisherNativeWrapMode.None) continue;
                if (page.IsMaster || mode != PublisherNativeWrapMode.RectangularBoth) return false;
                Bounds? obstacle = Placement(state, id);
                if (!obstacle.HasValue || wrap == null) return false;
                Bounds value = obstacle.Value; OfficeTextPadding margin = wrap.Margins;
                if (value.X - margin.Left < target.Value.X + target.Value.Width
                    && value.X + value.Width + margin.Right > target.Value.X
                    && value.Y - margin.Top < target.Value.Y + target.Value.Height
                    && value.Y + value.Height + margin.Bottom > target.Value.Y) ids.Add(id);
            }
            result.Add(frameId, Array.AsReadOnly(ids.ToArray()));
        }
        return true;
    }

    private static bool Authored(PublisherPublicationState state, uint id) =>
        state.Shapes.TryGetValue(id, out PublisherShape? shape) && shape.IsAuthored
        || state.Pictures.TryGetValue(id, out PublisherPicture? picture) && picture.IsAuthored
        || state.Text.Frames.TryGetValue(id, out PublisherFrameState? frame) && frame.IsAuthored
        || state.Tables.TryGetValue(id, out PublisherTableState? table) && table.IsAuthored;

    private static Bounds? Placement(PublisherPublicationState state, uint id) {
        if (state.Text.Frames.TryGetValue(id, out PublisherFrameState? frame))
            return frame.CanEditBounds && !frame.Hidden && frame.HasAnchor && frame.TransformAvailable
                ? new(frame.X, frame.Y, frame.Width, frame.Height) : null;
        if (state.Pictures.TryGetValue(id, out PublisherPicture? picture))
            return picture.CanEdit ? new(picture.X, picture.Y, picture.Width, picture.Height) : null;
        if (state.Shapes.TryGetValue(id, out PublisherShape? shape))
            return shape.CanEdit ? new(shape.X, shape.Y, shape.Width, shape.Height) : null;
        return null;
    }

    private readonly record struct Bounds(double X, double Y, double Width, double Height);
}

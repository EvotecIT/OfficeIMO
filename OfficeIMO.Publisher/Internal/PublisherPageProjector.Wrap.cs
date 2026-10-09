using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private IReadOnlyList<OfficeTextFlowRegion> TextExclusions(PublisherSourceShape frame, PublisherEscherShape source,
        FrameRectangle bounds, IReadOnlyDictionary<uint, PublisherEscherShape> artwork) {
        var result = new List<OfficeTextFlowRegion>();
        if (frame.WrapObjects.Count == 0) return result;
        OfficeTransform toFrame = PageTransform(source, bounds).Invert()
            .Then(OfficeTransform.Translate(-bounds.X, -bounds.Y));
        foreach (uint id in frame.WrapObjects) {
            _context.Record();
            if (!artwork.TryGetValue(id, out PublisherEscherShape? obstacle) || !obstacle.Bounds.HasValue
                || Owner(id, 0) != Owner(frame.Chunk.Id, 0)) {
                _context.Add("PUB_TEXT_WRAP_REFERENCE_UNRESOLVED", "A native text-wrap object has no resolved anchor on the frame's page. Its exclusion could not be applied.",
                    OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(frame.Chunk.Id));
                continue;
            }
            if (obstacle.Hidden) continue;
            FrameRectangle objectBounds = FrameBounds(obstacle, _source.Width, _source.Height);
            double left = WrapDistance(obstacle, 0x384), top = WrapDistance(obstacle, 0x385);
            double right = WrapDistance(obstacle, 0x386), bottom = WrapDistance(obstacle, 0x387);
            var local = PageTransform(obstacle, objectBounds).Then(toFrame)
                .TransformRectangleBounds(objectBounds.X - left, objectBounds.Y - top,
                    objectBounds.Width + left + right, objectBounds.Height + top + bottom);
            result.Add(new OfficeTextFlowRegion(local.Left, local.Top, local.Right - local.Left, local.Bottom - local.Top));
            if (obstacle.Properties.Any(property => property.PropertyId == 0x383))
                _context.Add("PUB_TIGHT_TEXT_WRAP_APPROXIMATED", "A native wrap polygon uses its transformed bounding rectangle. Tight and through-outline text paths are not reproduced.",
                    OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(id));
        }
        _context.Add("PUB_TEXT_WRAP_APPROXIMATED", "Native wrap references and distances exclude transformed object rectangles. Text uses the widest available interval per horizontal band; native side selection and outline wrapping can differ.",
            OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(frame.Chunk.Id));
        return result;
    }

    private double WrapDistance(PublisherEscherShape shape, int property) {
        int distance = unchecked((int)(shape.Property(property) ?? 0));
        if (distance < 0) _context.Add("PUB_NEGATIVE_WRAP_DISTANCE_APPROXIMATED", "A negative native wrap distance was bounded to the object's outline.",
            OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(shape.Id));
        return Math.Max(0, distance) / 12700D;
    }

    private static OfficeImageFrameTransform Transform(PublisherEscherShape shape, FrameRectangle bounds) =>
        new(shape.Transform.RotationDegrees.GetValueOrDefault(), bounds.X + bounds.Width / 2,
            bounds.Y + bounds.Height / 2, shape.Transform.FlipHorizontal, shape.Transform.FlipVertical);

    private OfficeTransform PageTransform(PublisherEscherShape shape, FrameRectangle bounds) =>
        Transform(shape, bounds).CreateDestinationTransform().Then(GroupPageTransform(shape));

    private OfficeTransform GroupPageTransform(PublisherEscherShape shape) =>
        shape.GroupTransform == OfficeTransform.Identity ? OfficeTransform.Identity :
        OfficeTransform.Translate(-_source.Width / 2, -_source.Height / 2)
            .Then(OfficeTransform.Scale(12700, 12700))
            .Then(shape.GroupTransform)
            .Then(OfficeTransform.Scale(1 / 12700D, 1 / 12700D))
            .Then(OfficeTransform.Translate(_source.Width / 2, _source.Height / 2));
}

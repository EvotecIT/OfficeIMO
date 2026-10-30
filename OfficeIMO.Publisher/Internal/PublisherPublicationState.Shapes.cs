using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPublicationState {
    /// <summary>Text-capable artwork reads its current placement from the frame owner, never from retained native coordinates.</summary>
    internal PublisherShape CurrentShape(uint id) {
        PublisherShape shape = Shapes[id];
        if (!Text.Frames.TryGetValue(id, out PublisherFrameState? frame)) return shape;
        bool placement = frame.HasAnchor && frame.TransformAvailable && frame.Width >= 0 && frame.Height >= 0
            && (frame.Width > 0 || frame.Height > 0);
        return shape.WithPlacement(frame.X, frame.Y, frame.Width, frame.Height, frame.PageTransform, placement);
    }

    internal PublisherShape EditableShape(uint id) {
        if (!Shapes.ContainsKey(id)) throw new ArgumentOutOfRangeException(nameof(id), "The publication has no such vector shape.");
        PublisherShape shape = CurrentShape(id);
        if (!shape.CanEdit) throw new NotSupportedException(shape.EditLimitation);
        return shape;
    }

    internal PublisherPublicationState SetShapeBounds(uint id, double x, double y, double width, double height) {
        PublisherShape shape = EditableShape(id);
        PublisherShapeGeometry.ValidateBounds(x, y, width, height, shape.CreateShape().Kind == OfficeShapeKind.Line);
        PublisherPageState page = Pages.Single(item => item.Id == shape.PageId);
        page.FrameEditRange.Validate(x, y, width, height);
        if (Text.Frames.ContainsKey(id)) return WithText(Text.SetBounds(id, x, y, width, height));
        var shapes = Shapes.ToDictionary(item => item.Key, item => item.Value);
        shapes[id] = shape.WithPlacement(x, y, width, height, OfficeTransform.Translate(x, y), true);
        return WithShapes(shapes);
    }

    internal PublisherPublicationState SetShapePaint(uint id, OfficeColor? fill, OfficeColor? stroke,
        double strokeWidth, double fillOpacity, double strokeOpacity) {
        PublisherShape shape = EditableShape(id);
        PublisherShapeGeometry.ValidatePaint(strokeWidth, fillOpacity, strokeOpacity);
        var shapes = Shapes.ToDictionary(item => item.Key, item => item.Value);
        shapes[id] = shape.WithPaint(fill, stroke, strokeWidth, fillOpacity, strokeOpacity);
        return WithShapes(shapes);
    }

    internal PublisherPublicationState AddShape(uint pageId, OfficeShape shape, double x, double y, out uint shapeId) {
        PublisherPageState? page = Pages.SingleOrDefault(item => item.Id == pageId);
        if (page == null) throw new ArgumentOutOfRangeException(nameof(pageId), "The publication has no such page.");
        if (!page.SupportsPlacementEditing) throw new NotSupportedException("This master page's coordinate registration is not qualified for shape creation.");
        page.FrameEditRange.Validate(x, y, shape.Width, shape.Height);
        if (NextObjectId == 0) throw new InvalidOperationException("The publication object identifier space is exhausted.");
        shapeId = NextObjectId;
        uint createdId = shapeId;
        var shapes = Shapes.ToDictionary(item => item.Key, item => item.Value);
        shapes.Add(createdId, new PublisherShape(createdId, pageId, x, y, shape.Width, shape.Height,
            OfficeTransform.Translate(x, y), true, shape, null, true));
        var owners = ObjectOwners.ToDictionary(item => item.Key, item => item.Value); owners.Add(createdId, pageId);
        return new PublisherPublicationState(Pages.Select(item => item.Id == pageId ? item.WithObject(createdId) : item),
            owners, Pictures, Images, Text, _sourceImageIds, NextImageId, createdId == uint.MaxValue ? 0 : createdId + 1, shapes);
    }

    private PublisherPublicationState WithShapes(IReadOnlyDictionary<uint, PublisherShape> shapes) =>
        new(Pages, ObjectOwners, Pictures, Images, Text, _sourceImageIds, NextImageId, NextObjectId, shapes);
}

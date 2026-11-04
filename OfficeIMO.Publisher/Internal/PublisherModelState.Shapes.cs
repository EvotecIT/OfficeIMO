using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherModelState {
    internal PublisherModelState SetShapeBounds(uint id, double x, double y, double width, double height, CancellationToken token) =>
        WithPublication(_publication.SetShapeBounds(id, x, y, width, height), token);

    internal PublisherModelState SetShapePaint(uint id, OfficeColor? fill, OfficeColor? stroke,
        double strokeWidth, double fillOpacity, double strokeOpacity) =>
        WithPublication(_publication.SetShapePaint(id, fill, stroke, strokeWidth, fillOpacity, strokeOpacity));

    internal PublisherModelState AddShape(uint pageId, OfficeShape shape, double x, double y,
        CancellationToken token, out uint shapeId) {
        token.ThrowIfCancellationRequested();
        if ((long)SourceObjectCount + _publication.AuthoredObjectCount >= _decoded.Options.Limits.MaxItems)
            throw new InvalidDataException("Publisher current object limit exceeded.");
        return WithPublication(_publication.AddShape(pageId, shape, x, y, out shapeId));
    }
}

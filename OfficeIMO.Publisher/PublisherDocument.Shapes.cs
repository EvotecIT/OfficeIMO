using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    /// <summary>Moves or resizes supported vector artwork using physical page-local bounds. Text-capable artwork shares its current placement with the text frame of the same ID.</summary>
    /// <remarks>Line extents may be zero on one axis; other extents must be positive. Finite negative offsets permit clipped bleed. Unsupported artwork is rejected. A successful edit recalculates text flow and every exported scene; failure leaves the publication unchanged. Existing snapshots remain detached and concurrent mutation is unsupported.</remarks>
    public void SetShapeBounds(uint shapeId, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable native model.");
        Commit(model.SetShapeBounds(shapeId, x.ToPoints(), y.ToPoints(), width.ToPoints(), height.ToPoints(), cancellationToken).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Replaces the supported shape's basic solid fill and outline. Null colors disable that paint; physical stroke width defaults to one point and opacities must be between zero and one.</summary>
    /// <remarks>Geometry, stroke caps and joins remain unchanged. Artwork with gradients, effects, custom geometry, transformations or unresolved paint remains inspection-only. Failed edits leave the publication unchanged; old snapshots remain detached.</remarks>
    public void SetShapePaint(uint shapeId, OfficeColor? fillColor, OfficeColor? strokeColor, OfficeLength? strokeWidth = null,
        double fillOpacity = 1, double strokeOpacity = 1, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable native model.");
        Commit(model.SetShapePaint(shapeId, fillColor, strokeColor, (strokeWidth ?? OfficeLength.Points(1)).ToPoints(), fillOpacity, strokeOpacity).Project(cancellationToken), cancellationToken);
    }

    /// <summary>Appends a rectangle, rounded rectangle, ellipse or nondegenerate line to an existing document or master page, returning its current immutable snapshot.</summary>
    /// <remarks>The origin uses physical lengths; the drawing-owned descriptor retains its local point coordinates. The descriptor is copied and must use solid paint without transforms, clipping, effects, dashes or markers. A line may have zero extent on one axis. New artwork has no story or wrap references. This edits the in-memory publication and does not provide native PUB creation or saving. Validation, cancellation or projection failure leaves the document and its identifier allocator unchanged.</remarks>
    public PublisherShape AddShape(uint pageId, OfficeShape shape, OfficeLength x, OfficeLength y, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(shape);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeShape owned = PublisherShapeGeometry.OwnedPrimitive(shape, x.ToPoints(), y.ToPoints());
        PublisherModelState model = _model ?? throw new NotSupportedException("This publication has no editable native model.");
        PublisherDocument candidate = model.AddShape(pageId, owned, x.ToPoints(), y.ToPoints(), cancellationToken, out uint id).Project(cancellationToken);
        PublisherShape result = candidate.Pages.Concat(candidate.MasterPages).SelectMany(page => page.Shapes).Single(item => item.Id == id);
        Commit(candidate, cancellationToken);
        return result;
    }
}

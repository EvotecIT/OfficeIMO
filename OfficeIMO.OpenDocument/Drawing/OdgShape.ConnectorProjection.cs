namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    // Projection overlays only attachment bounds. Public editing/routing getters
    // continue to resolve saved XML; no temporary source mutation is required.
    private IReadOnlyDictionary<XElement, OdfRect>? _projectionBounds;

    internal OdfRect AttachmentBounds => _projectionBounds != null &&
        _projectionBounds.TryGetValue(Element, out OdfRect bounds) ? bounds : Bounds;

    internal OdgShape WithProjectionBounds(IReadOnlyDictionary<XElement, OdfRect> bounds) =>
        new OdgShape((OdgDocument)Document, Element) { _projectionBounds = bounds };

    private OdgShape AttachmentShape(XElement element) =>
        new OdgShape((OdgDocument)Document, element) { _projectionBounds = _projectionBounds };
}

using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    // Track the visible page (or VisualBrush viewbox) back through native transforms.
    // Its conservative local bounding rectangle is sufficient for a user-space mask,
    // even when content outside the page is translated or rotated into view.
    private static BrushRegion LocalRegion(BrushRegion visible, string? transform) {
        if (transform == null) return visible;
        if (!OfficeSvgTransformParser.TryParse(transform, out var matrix)) throw new InvalidDataException("Invalid visual transform.");
        // A singular affine map has no painted area; retain a finite mask region.
        if (!matrix.TryInvert(out var inverse)) return visible;
        var corners = new[] {
            inverse.TransformPoint(new OfficePoint(visible.X, visible.Y)),
            inverse.TransformPoint(new OfficePoint(visible.X + visible.Width, visible.Y)),
            inverse.TransformPoint(new OfficePoint(visible.X, visible.Y + visible.Height)),
            inverse.TransformPoint(new OfficePoint(visible.X + visible.Width, visible.Y + visible.Height))
        };
        double x = corners.Min(p => p.X), y = corners.Min(p => p.Y);
        return new BrushRegion(x, y, corners.Max(p => p.X) - x, corners.Max(p => p.Y) - y);
    }
    private void ApplyOpacityMask(XElement source, XElement target, Dictionary<string, Resource> scope, string part, int depth, BrushRegion region) {
        if (source.Attribute("OpacityMask") == null && source.Element(source.Name.Namespace + (source.Name.LocalName + ".OpacityMask")) == null) return;
        Charge(depth);
        string id = "mask" + (++_id);
        var mask = Element("mask", new XAttribute("id", id), new XAttribute("mask-type", "alpha"), new XAttribute("maskUnits", "userSpaceOnUse"), new XAttribute("maskContentUnits", "userSpaceOnUse"),
            new XAttribute("x", N(region.X)), new XAttribute("y", N(region.Y)), new XAttribute("width", N(region.Width)), new XAttribute("height", N(region.Height)));
        var paint = Element("path", new XAttribute("d", "M" + N(region.X) + "," + N(region.Y) + " h" + N(region.Width) + " v" + N(region.Height) + " h" + N(-region.Width) + " Z"));
        Paint(source, "OpacityMask", paint, "fill", scope, part, depth);
        mask.Add(ApplyBrushFill(paint));
        _defs.Add(mask); Set(target, "mask", "url(#" + id + ")");
    }
}

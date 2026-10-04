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
        return TransformRegion(visible, inverse);
    }
    private static BrushRegion TransformRegion(BrushRegion visible, OfficeTransform matrix) {
        var corners = new[] {
            matrix.TransformPoint(new OfficePoint(visible.X, visible.Y)),
            matrix.TransformPoint(new OfficePoint(visible.X + visible.Width, visible.Y)),
            matrix.TransformPoint(new OfficePoint(visible.X, visible.Y + visible.Height)),
            matrix.TransformPoint(new OfficePoint(visible.X + visible.Width, visible.Y + visible.Height))
        };
        double x = corners.Min(p => p.X), y = corners.Min(p => p.Y);
        return new BrushRegion(x, y, corners.Max(p => p.X) - x, corners.Max(p => p.Y) - y);
    }
    private static BrushRegion IntersectRegion(BrushRegion a, BrushRegion b) {
        double x = Math.Max(a.X, b.X), y = Math.Max(a.Y, b.Y);
        return new BrushRegion(x, y, Math.Max(0, Math.Min(a.X + a.Width, b.X + b.Width) - x), Math.Max(0, Math.Min(a.Y + a.Height, b.Y + b.Height) - y));
    }
    private XElement ApplyCoverageMask(XElement coverage, XElement paint, BrushRegion region) {
        if (region.Width <= 0 || region.Height <= 0) return Element("g");
        string id = "coverageMask" + (++_id);
        _defs.Add(Element("mask", new XAttribute("id", id), new XAttribute("mask-type", "alpha"), new XAttribute("maskUnits", "userSpaceOnUse"), new XAttribute("maskContentUnits", "userSpaceOnUse"),
            new XAttribute("x", N(region.X)), new XAttribute("y", N(region.Y)), new XAttribute("width", N(region.Width)), new XAttribute("height", N(region.Height)), coverage));
        // A nested viewport also bounds Core's retained effect surfaces. A small
        // mask rectangle in a page-sized drawing alone would still allocate and
        // charge page-sized intermediate surfaces for every native run/stroke.
        return Element("svg", new XAttribute("x", N(region.X)), new XAttribute("y", N(region.Y)),
            new XAttribute("width", N(region.Width)), new XAttribute("height", N(region.Height)),
            new XAttribute("viewBox", string.Join(" ", new[] { region.X, region.Y, region.Width, region.Height }.Select(N))),
            Element("g", new XAttribute("mask", "url(#" + id + ")"), paint));
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

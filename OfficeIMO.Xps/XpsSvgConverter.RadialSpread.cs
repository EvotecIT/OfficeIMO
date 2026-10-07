using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private void NativeRadialField(XElement source, XElement gradient, XElement target, string attribute, BrushRegion region, string spread, BrushRegion? paintBounds) {
        bool objectBounds = (string?)gradient.Attribute("gradientUnits") == "objectBoundingBox";
        // target.d may contain only filled figures; stroke callers provide the
        // full path bounds for both coordinate mapping and cycle expansion.
        var bounds = paintBounds;
        if (!bounds.HasValue && (objectBounds || spread != "Pad"))
            StripFillRule((string?)target.Attribute("d") ?? "", out _, out bounds);
        if (objectBounds) {
            if (!bounds.HasValue || bounds.Value.Width <= 0 || bounds.Value.Height <= 0) { Set(target, attribute, "none"); return; }
            var box = bounds.Value;
            Set(gradient, "gradientTransform", "matrix(" + N(box.Width) + " 0 0 " + N(box.Height) + " " + N(box.X) + " " + N(box.Y) + ") " + (string?)gradient.Attribute("gradientTransform"));
        }
        Set(gradient, "gradientUnits", "userSpaceOnUse");
        if (spread != "Pad") {
            if (bounds.HasValue) {
                double margin = attribute == "stroke"
                    ? XpsPackage.Number((string?)source.Attribute("StrokeThickness"), 1) * Math.Max(1, XpsPackage.Number((string?)source.Attribute("StrokeMiterLimit"), 10)) : 0;
                var box = bounds.Value;
                region = IntersectRegion(region, new BrushRegion(box.X - margin, box.Y - margin, box.Width + 2 * margin, box.Height + 2 * margin));
            }
            if (!ExpandNativeRadialSpread(gradient, region, spread)) {
                // SVG 2 can repeat the reversed shrinking-circle field directly,
                // including infinitely many cycles at a tangent boundary. Keep
                // finite expansion when available for shared Drawing import.
                NativeRadialPad(gradient, target, attribute, region, spread == "Reflect");
                return;
            }
        }
        NativeRadialPad(gradient, target, attribute, region);
    }

    private bool ExpandNativeRadialSpread(XElement gradient, BrushRegion region, string spread) {
        if (spread != "Repeat" && spread != "Reflect") return false;
        double Value(string name) => XpsPackage.Number((string?)gradient.Attribute(name));
        double fx = Value("fx"), fy = Value("fy"), cx = Value("cx"), cy = Value("cy"), radius = Value("r");
        if (!OfficeSvgTransformParser.TryParse((string?)gradient.Attribute("gradientTransform"), out var transform) || !transform.TryInvert(out _)) return false;
        var field = new OfficeRadialGradient(fx, fy, 0, cx, cy, radius,
            new OfficeGradientStop(0, OfficeColor.Black), new OfficeGradientStop(1, OfficeColor.White)).TransformCoordinates(transform);
        var corners = new[] { new OfficePoint(region.X, region.Y), new OfficePoint(region.X + region.Width, region.Y),
            new OfficePoint(region.X + region.Width, region.Y + region.Height), new OfficePoint(region.X, region.Y + region.Height) };
        var stops = gradient.Elements().ToArray();
        const int stopBudget = 256;
        if (stops.Length == 0 || stops.Length > stopBudget ||
            !field.TryGetSpreadCycles(corners, true, spread == "Reflect", stopBudget, out int cycles) ||
            (long)cycles * (stops.Length + 2) > stopBudget) return false;
        double previous = 0;
        foreach (var stop in stops) {
            previous = Math.Max(previous, Math.Min(1, Math.Max(0, XpsPackage.Number((string?)stop.Attribute("offset")))));
            Set(stop, "offset", N(previous));
        }
        gradient.RemoveNodes();
        void Add(XElement stop, double position) {
            var copy = CloneProjection(stop); Set(copy, "offset", N(position / cycles)); gradient.Add(copy);
        }
        for (int cycle = 0; cycle < cycles; cycle++) {
            bool reverse = spread == "Reflect" && (cycle & 1) != 0;
            Add(reverse ? stops[stops.Length - 1] : stops[0], cycle);
            foreach (var stop in reverse ? Enumerable.Reverse(stops) : stops) {
                double offset = XpsPackage.Number((string?)stop.Attribute("offset"));
                Add(stop, cycle + (reverse ? 1 - offset : offset));
            }
            Add(reverse ? stops[0] : stops[stops.Length - 1], cycle + 1);
        }
        Set(gradient, "cx", N(fx + cycles * (cx - fx))); Set(gradient, "cy", N(fy + cycles * (cy - fy)));
        Set(gradient, "r", N(radius * cycles)); Set(gradient, "spreadMethod", "pad");
        return true;
    }
}

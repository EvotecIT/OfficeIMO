using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

internal sealed class XpsNavigationTarget {
    internal XpsNavigationTarget(string name, OfficePoint[] corners) { Name = name; Corners = corners; }
    internal string Name { get; }
    internal OfficePoint[] Corners { get; }
    internal OfficePoint TopLeft => new(Corners.Min(p => p.X), Corners.Min(p => p.Y));
}

internal sealed partial class XpsSvgConverter {
    private readonly Dictionary<XElement, BrushRegion> _nativeBounds = new();
    private readonly List<XpsNavigationTarget> _targets = new();

    private void RecordNavigationBounds(XElement child, XElement parent, string? value, int firstTarget) {
        if (_visualDepth != 0) return;
        BrushRegion bounds = _nativeBounds.TryGetValue(child, out var known) ? known : new BrushRegion(0, 0, 0, 0);
        if (child.Attribute("Name") is XAttribute name) _targets.Add(new XpsNavigationTarget(name.Value, new[] {
            new OfficePoint(bounds.X, bounds.Y), new OfficePoint(bounds.X + bounds.Width, bounds.Y),
            new OfficePoint(bounds.X + bounds.Width, bounds.Y + bounds.Height), new OfficePoint(bounds.X, bounds.Y + bounds.Height)
        }));
        if (value != null) {
            if (!OfficeSvgTransformParser.TryParse(value, out var transform)) throw new InvalidDataException("Invalid navigation transform.");
            bounds = TransformRegion(bounds, transform);
            for (int i = firstTarget; i < _targets.Count; i++) {
                _token.ThrowIfCancellationRequested();
                var corners = _targets[i].Corners;
                for (int corner = 0; corner < corners.Length; corner++) corners[corner] = transform.TransformPoint(corners[corner]);
            }
        }
        if (_nativeBounds.TryGetValue(parent, out var previous)) {
            double x = Math.Min(previous.X, bounds.X), y = Math.Min(previous.Y, bounds.Y);
            bounds = new BrushRegion(x, y, Math.Max(previous.X + previous.Width, bounds.X + bounds.Width) - x,
                Math.Max(previous.Y + previous.Height, bounds.Y + bounds.Height) - y);
        }
        _nativeBounds[parent] = bounds;
    }
}

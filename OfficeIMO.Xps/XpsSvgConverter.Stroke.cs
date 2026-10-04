using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private void Stroke(XElement source, XElement target, string path) {
        double thickness = XpsPackage.Number((string?)source.Attribute("StrokeThickness"), 1);
        double miterLimit = XpsPackage.Number((string?)source.Attribute("StrokeMiterLimit"), 10);
        if (thickness < 0 || miterLimit < 1) throw new InvalidDataException("Invalid XPS stroke dimensions.");
        Set(target, "stroke-width", N(thickness));
        Set(target, "stroke-miterlimit", N(miterLimit));
        bool hasStroke = (string?)target.Attribute("stroke") != "none" && thickness > 0;
        double[] dashes = Numbers((string?)source.Attribute("StrokeDashArray") ?? "");
        if (dashes.Any(n => n < 0)) throw new InvalidDataException("Negative XPS stroke dash length.");
        if (dashes.Length > 0) Set(target, "stroke-dasharray", string.Join(" ", dashes.Select(n => N(n * thickness))));
        if (source.Attribute("StrokeDashOffset") is XAttribute offset) Set(target, "stroke-dashoffset", N(XpsPackage.Number(offset.Value) * thickness));
        string cap = (string?)source.Attribute("StrokeStartLineCap") ?? "Flat";
        string endCap = (string?)source.Attribute("StrokeEndLineCap") ?? "Flat";
        string dashCap = (string?)source.Attribute("StrokeDashCap") ?? "Flat";
        if (hasStroke) {
            if (cap != endCap) Loss("Unequal stroke end caps");
            if (dashes.Length > 0 && dashCap != cap) Loss("Separate stroke dash caps");
            if (cap == "Triangle" || endCap == "Triangle" || (dashes.Length > 0 && dashCap == "Triangle")) Loss("Triangle stroke cap");
        }
        if (!new[] { "Flat", "Round", "Square", "Triangle" }.Contains(cap) || !new[] { "Flat", "Round", "Square", "Triangle" }.Contains(endCap) || !new[] { "Flat", "Round", "Square", "Triangle" }.Contains(dashCap)) throw new InvalidDataException("Invalid XPS stroke cap.");
        Set(target, "stroke-linecap", cap == "Round" ? "round" : cap == "Square" ? "square" : "butt");
        string join = (string?)source.Attribute("StrokeLineJoin") ?? "Miter";
        if (!new[] { "Miter", "Round", "Bevel" }.Contains(join)) throw new InvalidDataException("Invalid XPS stroke join.");
        Set(target, "stroke-linejoin", join.ToLowerInvariant());
        if (hasStroke && join == "Miter" && RequiresClippedMiter(path, miterLimit)) Loss("Clipped miter stroke join");
    }

    // SVG bevels over-limit miters; XPS clips their tips. Detect those joins using the
    // actual line/Bezier endpoint tangents instead of silently substituting the bevel.
    private static bool RequiresClippedMiter(string path, double limit) {
        if (!OfficeSvgPathDataParser.TryParse(path, 100000, out var commands, out _, allowEmptyGeometry: true)) throw new InvalidDataException("Invalid stroke path.");
        OfficePoint current = default, start = default, incoming = default, first = default;
        bool hasTangent = false, pendingDegenerate = false, leadingDegenerate = false;
        OfficePoint Difference(OfficePoint a, OfficePoint b) => new(a.X - b.X, a.Y - b.Y);
        bool Zero(OfficePoint p) => p.X == 0 && p.Y == 0;
        bool OverLimit(OfficePoint a, OfficePoint b, double joinLimit) {
            double dot = (a.X * b.X + a.Y * b.Y) / Math.Sqrt((a.X * a.X + a.Y * a.Y) * (b.X * b.X + b.Y * b.Y));
            return 1 + Math.Max(-1, Math.Min(1, dot)) < 2 / (joinLimit * joinLimit) - 1e-12;
        }
        foreach (var command in commands) {
            if (command.Kind == OfficePathCommandKind.MoveTo) { current = start = command.Point; hasTangent = pendingDegenerate = leadingDegenerate = false; continue; }
            bool closed = command.Kind == OfficePathCommandKind.Close;
            OfficePoint end = closed ? start : command.Point;
            OfficePoint outgoing = Difference(end, current), ending = outgoing;
            if (command.Kind == OfficePathCommandKind.CubicBezierTo) {
                outgoing = Difference(command.ControlPoint1, current);
                if (Zero(outgoing)) outgoing = Difference(command.ControlPoint2, current);
                if (Zero(outgoing)) outgoing = Difference(end, current);
                ending = Difference(end, command.ControlPoint2);
                if (Zero(ending)) ending = Difference(end, command.ControlPoint1);
                if (Zero(ending)) ending = Difference(end, current);
            } else if (command.Kind == OfficePathCommandKind.QuadraticBezierTo) {
                outgoing = Difference(command.ControlPoint1, current);
                if (Zero(outgoing)) outgoing = Difference(end, current);
                ending = Difference(end, command.ControlPoint1);
                if (Zero(ending)) ending = Difference(end, current);
            }
            if (!Zero(outgoing) && !Zero(ending)) {
                if (hasTangent && OverLimit(incoming, outgoing, pendingDegenerate ? 1 : limit)) return true;
                if (!hasTangent) first = outgoing;
                incoming = ending; hasTangent = true; pendingDegenerate = false;
            } else if (!closed) {
                // ECMA-388 18.6.7.3 uses an implied miter limit of one across
                // degenerate segments, including at the seam of closed figures.
                pendingDegenerate = true;
                if (!hasTangent) leadingDegenerate = true;
            }
            if (closed && hasTangent && OverLimit(incoming, first, pendingDegenerate || leadingDegenerate ? 1 : limit)) return true;
            current = end;
        }
        return false;
    }
}

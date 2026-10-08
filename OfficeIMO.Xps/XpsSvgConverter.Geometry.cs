using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private sealed class StrokeFigure {
        internal StrokeFigure(string path, bool start, bool end) { Path = path; UseStartCap = start; UseEndCap = end; }
        internal string Path { get; }
        internal bool UseStartCap { get; }
        internal bool UseEndCap { get; }
    }
    private sealed class GeometryProjection {
        internal string Full { get; set; } = "";
        internal string Fill { get; set; } = "";
        internal List<StrokeFigure> Strokes { get; } = new();
        internal bool HasSeparateStrokeGeometry { get; set; }
    }

    private GeometryProjection? PathProjection(XElement path, Dictionary<string, Resource> scope) {
        XElement? geometry = path.Element(path.Name.Namespace + "Path.Data")?.Elements().SingleOrDefault();
        string? value = (string?)path.Attribute("Data");
        if (value?.StartsWith("{", StringComparison.Ordinal) == true) geometry = ResolveResource(value, scope)!.Value;
        if (geometry?.Name.LocalName != "PathGeometry" || geometry.Attribute("Figures") != null) return null;
        CheckAttributes(geometry, "Figures FillRule");
        var projection = ProjectFigures(geometry);
        projection.Full = ((string?)geometry.Attribute("FillRule") == "NonZero" ? "F1 " : "F0 ") + projection.Full;
        return projection;
    }

    private GeometryProjection ProjectFigures(XElement geometry) {
        var projection = new GeometryProjection();
        var full = new StringBuilder(); var fill = new StringBuilder();
        foreach (var figure in geometry.Elements()) {
            Charge(1);
            if (figure.Name.LocalName != "PathFigure") { Loss("Geometry child: " + figure.Name.LocalName); continue; }
            CheckAttributes(figure, "StartPoint IsClosed IsFilled");
            string start = (string?)figure.Attribute("StartPoint") ?? throw new InvalidDataException("Missing figure start.");
            bool closed = NativeBoolean(figure, "IsClosed", false), filled = NativeBoolean(figure, "IsFilled", true);
            projection.HasSeparateStrokeGeometry |= !filled;
            var segments = new List<(string Command, string End, bool Stroked)>();
            foreach (var segment in figure.Elements()) {
                Charge(1);
                CheckAttributes(segment, "Point Points Point1 Point2 Point3 Size RotationAngle IsLargeArc SweepDirection IsStroked");
                string command, end;
                if (segment.Name.LocalName == "ArcSegment") {
                    end = (string?)segment.Attribute("Point") ?? throw new InvalidDataException("Missing arc endpoint.");
                    command = "A " + (string?)segment.Attribute("Size") + " " + ((string?)segment.Attribute("RotationAngle") ?? "0") + " " +
                        (NativeBoolean(segment, "IsLargeArc", false) ? "1" : "0") + " " +
                        ((string?)segment.Attribute("SweepDirection") == "Clockwise" ? "1" : "0") + " " + end;
                } else {
                    string? code = segment.Name.LocalName == "PolyLineSegment" ? "L" :
                        segment.Name.LocalName == "PolyBezierSegment" ? "C" : segment.Name.LocalName == "PolyQuadraticBezierSegment" ? "Q" : null;
                    if (code == null) { Loss("Path segment: " + segment.Name.LocalName); continue; }
                    string points = (string?)segment.Attribute("Points") ?? throw new InvalidDataException("Missing segment points.");
                    var values = Numbers(points);
                    if (values.Length < 2) throw new InvalidDataException("Missing segment endpoint.");
                    end = N(values[values.Length - 2]) + "," + N(values[values.Length - 1]);
                    command = code + " " + points;
                }
                bool stroked = NativeBoolean(segment, "IsStroked", true);
                projection.HasSeparateStrokeGeometry |= !stroked;
                segments.Add((command, end, stroked));
            }
            string figureStart = "M " + start;
            string complete = figureStart + " " + string.Join(" ", segments.Select(s => s.Command)) + (closed ? " Z " : " ");
            full.Append(complete);
            if (filled) fill.Append(complete);
            if (!segments.Any(s => !s.Stroked)) { projection.Strokes.Add(new StrokeFigure(complete, !closed, !closed)); continue; }
            if (closed) segments.Add(("L " + start, start, true));
            var runs = new List<StrokeFigure>(); StringBuilder? run = null;
            string current = start; bool startCap = false;
            for (int i = 0; i < segments.Count; i++) {
                var segment = segments[i];
                if (segment.Stroked) {
                    if (run == null) { run = new StringBuilder("M " + current + " "); startCap = !closed && i == 0; }
                    run.Append(segment.Command).Append(' ');
                } else if (run != null) { runs.Add(new StrokeFigure(run.ToString(), startCap, false)); run = null; }
                current = segment.End;
            }
            if (run != null) runs.Add(new StrokeFigure(run.ToString(), startCap, !closed));
            // A closed figure remains continuous across its authored seam when both
            // adjacent segments are stroked. Its first gap supplies the dash origin.
            if (closed && runs.Count > 1 && segments[0].Stroked) {
                var last = runs[runs.Count - 1];
                runs[runs.Count - 1] = new StrokeFigure(last.Path + runs[0].Path.Substring(figureStart.Length), false, false);
                runs.RemoveAt(0);
            }
            projection.Strokes.AddRange(runs);
        }
        projection.Full = full.ToString(); projection.Fill = fill.ToString(); return projection;
    }

    private static bool NativeBoolean(XElement element, string name, bool fallback) {
        string? value = (string?)element.Attribute(name);
        if (value == null) return fallback;
        if (value == "true" || value == "1") return true;
        if (value == "false" || value == "0") return false;
        throw new InvalidDataException("Invalid XPS boolean: " + name + ".");
    }

    private static bool HasDegenerateStrokeFigure(string path) {
        if (!OfficeSvgPathDataParser.TryParse(path, 100000, out var commands, out _, allowEmptyGeometry: true))
            throw new InvalidDataException("Invalid XPS path geometry.");
        OfficePoint start = default; bool degenerate = true, hasSegment = false;
        foreach (var command in commands) {
            if (command.Kind == OfficePathCommandKind.MoveTo) {
                if (degenerate && hasSegment) return true;
                start = command.Point; degenerate = true; hasSegment = false;
            }
            else if (command.Kind == OfficePathCommandKind.Close) { if (degenerate) return true; }
            else {
                hasSegment = true;
                if (!command.Point.Equals(start) ||
                (command.Kind == OfficePathCommandKind.CubicBezierTo && (!command.ControlPoint1.Equals(start) || !command.ControlPoint2.Equals(start))) ||
                (command.Kind == OfficePathCommandKind.QuadraticBezierTo && !command.ControlPoint1.Equals(start))) degenerate = false;
            }
        }
        return degenerate && hasSegment;
    }
}

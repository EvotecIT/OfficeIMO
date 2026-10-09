using OfficeIMO.Drawing;
using System.Text.RegularExpressions;

namespace OfficeIMO.OpenDocument;

/// <summary>Reads the bounded literal, single-paint-set enhanced-path profile without evaluating ODF equations.</summary>
internal static class OdfEnhancedGeometry {
    private const int MaximumTextLength = 1024 * 1024;
    private static readonly Regex TokenPattern = new Regex(@"\G[\s,]*(?:(?<number>[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)|(?<command>[A-Za-z]))", RegexOptions.CultureInvariant);

    internal static OfficeShape Project(XElement shape, double width, double height) {
        XElement geometry = shape.Element(OdfNamespaces.Draw + "enhanced-geometry")
            ?? throw new NotSupportedException("Custom-shape projection requires explicit enhanced geometry.");
        if (geometry.Attribute(OdfNamespaces.Draw + "path-stretchpoint-x") != null || geometry.Attribute(OdfNamespaces.Draw + "path-stretchpoint-y") != null)
            throw new NotSupportedException("Enhanced path stretch points are preserved but are not projected.");
        if (ReadBoolean(geometry, "extrusion")) throw new NotSupportedException("Extruded custom shapes are preserved but are not projected.");
        OdfViewBox box = OdfViewBox.Parse((string?)geometry.Attribute(OdfNamespaces.Svg + "viewBox"));
        string data = (string?)geometry.Attribute(OdfNamespaces.Draw + "enhanced-path")
            ?? throw new NotSupportedException("Custom-shape projection requires an explicit enhanced path; named presets are not substituted.");
        IReadOnlyList<Token> tokens = ReadTokens(data);
        if (tokens.Count == 0 || tokens[tokens.Count - 1].Command != 'N' || tokens.Count(token => token.Command == 'N') != 1)
            throw new NotSupportedException("Enhanced projection requires one paint set terminated by N; multiple independent paint sets are preserved but are not projected.");
        IReadOnlyList<OfficePathCommand> path;
        if (tokens.Any(token => token.Command == 'U')) path = FullEllipse(tokens);
        else {
            if (tokens.Any(token => token.Command != '\0' && token.Command is not ('M' or 'L' or 'C' or 'Q' or 'Z' or 'N')))
                throw new NotSupportedException("Enhanced projection supports literal M, L, C, Q and Z commands and a full U ellipse; other commands are preserved but are not projected.");
            path = OdgShape.ParsePath(string.Join(" ", tokens.Take(tokens.Count - 1).Select(token => token.ToString())));
        }
        bool mirrorX = ReadBoolean(geometry, "mirror-horizontal"), mirrorY = ReadBoolean(geometry, "mirror-vertical");
        foreach (OfficePathCommand command in path) {
            if (command.Kind != OfficePathCommandKind.Close) RequireInside(command.Point);
            if (command.Kind is OfficePathCommandKind.CubicBezierTo or OfficePathCommandKind.QuadraticBezierTo) RequireInside(command.ControlPoint1);
            if (command.Kind == OfficePathCommandKind.CubicBezierTo) RequireInside(command.ControlPoint2);
        }
        return OfficeShape.Path(Math.Max(width, .001), Math.Max(height, .001),
            OdfViewBoxProjection.Project(path, box, width, height, mirrorX, mirrorY));
        void RequireInside(OfficePoint point) {
            if (point.X < box.X || point.Y < box.Y || point.X > (double)box.X + box.Width || point.Y > (double)box.Y + box.Height)
                throw new NotSupportedException("Enhanced geometry with points or curve controls outside its view box is preserved but is outside the shared export profile.");
        }
    }

    private static IReadOnlyList<OfficePathCommand> FullEllipse(IReadOnlyList<Token> tokens) {
        if (tokens.Count is not (8 or 9) || tokens[0].Command != 'U' || tokens.Skip(1).Take(6).Any(token => token.Command != '\0') ||
            tokens.Count == 9 && tokens[7].Command != 'Z')
            throw new NotSupportedException("The supported U profile is one full ellipse, optionally closed, followed by N.");
        double centerX = tokens[1].Number, centerY = tokens[2].Number, rx = tokens[3].Number, ry = tokens[4].Number;
        double startAngle = tokens[5].Number, endAngle = tokens[6].Number;
        if (rx <= 0 || ry <= 0 || Math.Abs(endAngle - startAngle) != 360)
            throw new NotSupportedException("Enhanced U projection requires positive ellipse radii and exactly one full turn; partial arcs remain preserved.");
        double radialAngle = (startAngle % 360) * Math.PI / 180;
        double parametricAngle = Math.Atan2(rx * Math.Sin(radialAngle), ry * Math.Cos(radialAngle));
        if (parametricAngle < 0) parametricAngle += 2 * Math.PI;
        var start = PointAt(parametricAngle);
        var result = new List<OfficePathCommand> { OfficePathCommand.MoveTo(start) };
        // ODF U draws clockwise even when the full-turn end angle is smaller than the start angle.
        // Split at ellipse axes so controls fit the declared canvas, including non-cardinal starting rays.
        double angle = parametricAngle, remaining = 2 * Math.PI;
        while (remaining > 1e-12) {
            double nextAxis = (Math.Floor(angle / (Math.PI / 2)) + 1) * (Math.PI / 2);
            double sweep = Math.Min(remaining, nextAxis - angle);
            foreach (OfficePathCommand command in OfficeGeometry.CreateEllipticalArcCubicBezierCommands(PointAt(angle), rx, ry, angle, sweep))
                result.Add(OfficePathCommand.CubicBezierTo(SnapPoint(command.ControlPoint1), SnapPoint(command.ControlPoint2), SnapPoint(command.Point)));
            angle += sweep; remaining -= sweep;
        }
        if (tokens.Count == 9) result.Add(OfficePathCommand.Close());
        return result;
        OfficePoint PointAt(double value) => SnapPoint(new OfficePoint(centerX + rx * Math.Cos(value), centerY + ry * Math.Sin(value)));
        OfficePoint SnapPoint(OfficePoint point) => new OfficePoint(Snap(point.X, centerX - rx, centerX + rx), Snap(point.Y, centerY - ry, centerY + ry));
        double Snap(double value, double minimum, double maximum) {
            double tolerance = 1e-12 * Math.Max(1, Math.Max(Math.Abs(minimum), Math.Abs(maximum)));
            return Math.Abs(value - minimum) <= tolerance ? minimum : Math.Abs(value - maximum) <= tolerance ? maximum : value;
        }
    }

    private static IReadOnlyList<Token> ReadTokens(string data) {
        if (data.Length > MaximumTextLength) throw new InvalidDataException("Enhanced path data exceeds 1 MiB.");
        if (data.IndexOf('?') >= 0 || data.IndexOf('$') >= 0)
            throw new NotSupportedException("Enhanced path equations and modifier references are preserved but are not evaluated.");
        var tokens = new List<Token>();
        int position = 0;
        while (position < data.Length) {
            Match match = TokenPattern.Match(data, position);
            if (!match.Success) {
                if (data.Substring(position).All(character => char.IsWhiteSpace(character) || character == ',')) break;
                throw new FormatException("Invalid enhanced path notation.");
            }
            // An explicit cubic command has one command token and six coordinates.
            // The shared parser subsequently enforces the normalized command count.
            if (tokens.Count >= OdgShape.MaximumGeometryItems * 7 + 1) throw new InvalidDataException("Enhanced path data exceeds the geometry command limit.");
            if (match.Groups["command"].Success) tokens.Add(new Token(match.Groups["command"].Value[0], 0));
            else {
                double number = double.Parse(match.Groups["number"].Value, NumberStyles.Float, CultureInfo.InvariantCulture);
                if (double.IsNaN(number) || double.IsInfinity(number)) throw new InvalidDataException("Enhanced path coordinates must be finite.");
                tokens.Add(new Token('\0', number));
            }
            position = match.Index + match.Length;
        }
        return tokens;
    }

    private static bool ReadBoolean(XElement geometry, string name) => (string?)geometry.Attribute(OdfNamespaces.Draw + name) switch {
        null or "false" or "0" => false,
        "true" or "1" => true,
        _ => throw new FormatException("Invalid enhanced geometry boolean: " + name + ".")
    };

    private readonly struct Token {
        internal Token(char command, double number) { Command = command; Number = number; }
        internal char Command { get; }
        internal double Number { get; }
        public override string ToString() => Command == '\0' ? Number.ToString("R", CultureInfo.InvariantCulture) : Command.ToString();
    }
}

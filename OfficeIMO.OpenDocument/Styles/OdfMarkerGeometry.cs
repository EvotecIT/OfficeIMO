using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Immutable native marker geometry, shared by all shapes referencing a named marker.</summary>
public sealed class OdfMarkerGeometry {
    /// <summary>Creates bounded SVG path geometry. Its actual geometry bounds determine marker sizing and alignment.</summary>
    public OdfMarkerGeometry(OdfViewBox viewBox, string pathData) : this(CheckedViewBox(viewBox), pathData) { }
    /// <summary>Creates marker geometry with a native four-number view box, including decimal producer coordinates.</summary>
    public OdfMarkerGeometry(string viewBox, string pathData) {
        AuthoringViewBox = ValidateViewBox(viewBox);
        Commands = OdgShape.ParsePath(pathData);
        Bounds = OfficePathGeometry.Bounds(Commands);
        double width = Bounds.Right - Bounds.Left, height = Bounds.Bottom - Bounds.Top;
        if (!Finite(width) || !Finite(height) || width <= 0 || height <= 0)
            throw new ArgumentException("Marker geometry must occupy a finite two-dimensional area.", nameof(pathData));
        ViewBox = viewBox.Trim(); PathData = pathData;
    }
    /// <summary>Native SVG coordinate canvas. Creating or replacing a definition normalizes it to enclosing integer coordinates.</summary>
    public string ViewBox { get; }
    /// <summary>Original bounded path notation; curves and multiple contours are retained.</summary>
    public string PathData { get; }
    internal IReadOnlyList<OfficePathCommand> Commands { get; }
    internal (double Left, double Top, double Right, double Bottom) Bounds { get; }
    internal string AuthoringViewBox { get; }
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private static string CheckedViewBox(OdfViewBox value) { value.Validate(); return value.ToString(); }
    private static string ValidateViewBox(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (value.Length > 128) throw new ArgumentException("Marker view box exceeds the supported length.", nameof(value));
        string[] tokens = value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length != 4) throw new ArgumentException("Marker view box requires four numbers.", nameof(value));
        var numbers = new double[4];
        for (int i = 0; i < tokens.Length; i++)
            if (!double.TryParse(tokens[i], NumberStyles.Float, CultureInfo.InvariantCulture, out numbers[i]) || !Finite(numbers[i]) || (i >= 2 && numbers[i] <= 0))
                throw new ArgumentException("Marker view box requires finite coordinates and positive dimensions.", nameof(value));
        // ODF view boxes require integers, even though native producers sometimes emit fractional coordinates.
        double x = Math.Floor(numbers[0]), y = Math.Floor(numbers[1]);
        double width = Math.Ceiling(numbers[0] + numbers[2]) - x, height = Math.Ceiling(numbers[1] + numbers[3]) - y;
        if (x < int.MinValue || x > int.MaxValue || y < int.MinValue || y > int.MaxValue || !Finite(width) || !Finite(height) || width < 1 || width > int.MaxValue || height < 1 || height > int.MaxValue)
            throw new ArgumentException("Marker view box exceeds the supported integer canvas.", nameof(value));
        return new OdfViewBox((int)x, (int)y, (int)width, (int)height).ToString();
    }
}

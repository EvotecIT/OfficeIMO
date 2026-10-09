namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed named native gradient shared by referencing shapes.</summary>
public sealed class OdfGradient {
    private static readonly XNamespace LoExt = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
    private readonly OdfDocument _document;
    private readonly XElement _element;
    internal OdfGradient(OdfDocument document, XElement element) { _document = document; _element = element; }
    /// <summary>Name used by <see cref="OdfShape.FillGradientName"/>.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Draw + "name") ?? string.Empty;
    /// <summary>Reads or replaces a two-color native definition, retaining unrelated attributes.</summary>
    /// <remarks>Equivalent LibreOffice endpoint stops are retained and updated with the colors. SVG gradients and other stop extensions remain preserved but cannot be overwritten through this two-color property.</remarks>
    public OdfGradientPattern Pattern {
        get {
            try {
                EnsureTwoColorDefinition();
                OdfGradientStyle style = Read("style") switch {
                    "linear" => OdfGradientStyle.Linear, "axial" => OdfGradientStyle.Axial, "radial" => OdfGradientStyle.Radial,
                    "ellipsoid" => OdfGradientStyle.Ellipsoid, "square" => OdfGradientStyle.Square, "rectangular" => OdfGradientStyle.Rectangular,
                    _ => throw new InvalidDataException("Invalid native gradient style.")
                };
                if (style != OdfGradientStyle.Radial && HasAmbiguousLegacyAngle)
                    throw new NotSupportedException("A unitless gradient angle in ODF 1.2 may use producer-specific tenths of degrees; it is preserved without guessing.");
                return new OdfGradientPattern(style, OdfColor.Parse(Required("start-color")), OdfColor.Parse(Required("end-color")),
                    ReadAngle(), ReadPercent("border", 0), ReadPercent("start-intensity", 1), ReadPercent("end-intensity", 1),
                    ReadPercent("cx", 0), ReadPercent("cy", 0));
            } catch (Exception exception) when (exception is ArgumentException || exception is FormatException) {
                throw new InvalidDataException("Invalid gradient definition '" + Name + "'.", exception);
            }
        }
        set {
            if (value == null) throw new ArgumentNullException(nameof(value));
            EnsureTwoColorDefinition();
            WritePattern(_element, value);
            XElement[] stops = _element.Elements().ToArray();
            if (stops.Length == 2) {
                stops[0].SetAttributeValue(LoExt + "color-value", value.StartColor.ToString());
                stops[1].SetAttributeValue(LoExt + "color-value", value.EndColor.ToString());
            }
            _document.MarkPartDirty("styles.xml");
        }
    }
    private bool HasAmbiguousLegacyAngle => _document.Version == OdfVersion.V1_2 &&
        double.TryParse(Read("angle"), NumberStyles.Float, CultureInfo.InvariantCulture, out double angle) && angle != 0;
    internal static void WritePattern(XElement element, OdfGradientPattern pattern) {
        element.SetAttributeValue(OdfNamespaces.Draw + "style", pattern.Style.ToString().ToLowerInvariant());
        element.SetAttributeValue(OdfNamespaces.Draw + "start-color", pattern.StartColor.ToString());
        element.SetAttributeValue(OdfNamespaces.Draw + "end-color", pattern.EndColor.ToString());
        element.SetAttributeValue(OdfNamespaces.Draw + "angle", pattern.AngleDegrees.ToString("R", CultureInfo.InvariantCulture) + "deg");
        WritePercent("border", pattern.Border); WritePercent("start-intensity", pattern.StartIntensity); WritePercent("end-intensity", pattern.EndIntensity);
        WritePercent("cx", pattern.CenterX); WritePercent("cy", pattern.CenterY);
        void WritePercent(string name, double value) => element.SetAttributeValue(OdfNamespaces.Draw + name,
            FormatPercent(value));
    }
    // ODF percentages use decimal notation, so expand the round-trip double's exponent.
    private static string FormatPercent(double value) {
        string text = (value * 100).ToString("R", CultureInfo.InvariantCulture);
        int exponentAt = text.IndexOf('E');
        if (exponentAt < 0) return text + "%";
        int exponent = int.Parse(text.Substring(exponentAt + 1), NumberStyles.Integer, CultureInfo.InvariantCulture);
        string mantissa = text.Substring(0, exponentAt), sign = "";
        if (mantissa.StartsWith("-", StringComparison.Ordinal)) { sign = "-"; mantissa = mantissa.Substring(1); }
        int dot = mantissa.IndexOf('.');
        int position = (dot < 0 ? mantissa.Length : dot) + exponent;
        string digits = mantissa.Replace(".", "");
        return sign + (position <= 0 ? "0." + new string('0', -position) + digits :
            position >= digits.Length ? digits + new string('0', position - digits.Length) : digits.Insert(position, ".")) + "%";
    }
    private void EnsureTwoColorDefinition() {
        if (_element.Name != OdfNamespaces.Draw + "gradient")
            throw new NotSupportedException("Gradient '" + Name + "' uses SVG or extended gradient stops outside the native two-color model.");
        XElement[] children = _element.Elements().ToArray();
        if (children.Length == 0) return;
        if (children.Length != 2 || !EquivalentStop(children[0], 0, Required("start-color")) || !EquivalentStop(children[1], 1, Required("end-color")))
            throw new NotSupportedException("Gradient '" + Name + "' has extended stops that are not equivalent to its native endpoint colors.");
    }
    private static bool EquivalentStop(XElement element, double offset, string color) =>
        element.Name == LoExt + "gradient-stop" && !element.HasElements && (string?)element.Attribute(LoExt + "color-type") == "rgb" &&
        double.TryParse((string?)element.Attribute(OdfNamespaces.Svg + "offset"), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) && value == offset &&
        OdfColor.TryParse((string?)element.Attribute(LoExt + "color-value"), out OdfColor actual) && actual.Equals(OdfColor.Parse(color));
    private string? Read(string name) => (string?)_element.Attribute(OdfNamespaces.Draw + name);
    private string Required(string name) => Read(name) ?? throw new InvalidDataException("Gradient '" + Name + "' omits " + name + ".");
    private double ReadAngle() {
        string text = (Read("angle") ?? "0").Trim(); double factor = 1;
        if (text.EndsWith("grad", StringComparison.Ordinal)) { text = text.Substring(0, text.Length - 4); factor = 0.9; }
        else if (text.EndsWith("deg", StringComparison.Ordinal)) text = text.Substring(0, text.Length - 3);
        else if (text.EndsWith("rad", StringComparison.Ordinal)) { text = text.Substring(0, text.Length - 3); factor = 180 / Math.PI; }
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) || double.IsNaN(value) || double.IsInfinity(value * factor))
            throw new InvalidDataException("Invalid gradient angle.");
        return value * factor;
    }
    private double ReadPercent(string name, double fallback) {
        string? value = Read(name); if (value == null) return fallback;
        string text = value.Trim();
        if (!text.EndsWith("%", StringComparison.Ordinal)) throw new InvalidDataException("Gradient " + name + " requires a percentage.");
        text = text.Substring(0, text.Length - 1);
        if (text.Length == 0 || text.Any(character => character != '-' && character != '.' && (character < '0' || character > '9')) ||
            !double.TryParse(text, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out double number))
            throw new InvalidDataException("Invalid gradient percentage.");
        return number / 100;
    }
}

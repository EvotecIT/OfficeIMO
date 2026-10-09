namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed named dash definition shared by referencing shapes.</summary>
public sealed class OdfStrokeDash {
    private readonly OdfDocument _document;
    private readonly XElement _element;
    internal OdfStrokeDash(OdfDocument document, XElement element) { _document = document; _element = element; }
    /// <summary>Name used by <see cref="OdfShape.StrokeDashName"/>.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Draw + "name") ?? string.Empty;
    /// <summary>Reads or atomically replaces the supported pattern while retaining other definition metadata.</summary>
    /// <remarks>Imported definitions with omitted metrics or unsupported counts remain preserved but cannot be projected through this profile.</remarks>
    public OdfStrokeDashPattern Pattern {
        get {
            try {
                string? style = (string?)_element.Attribute(OdfNamespaces.Draw + "style");
                if (style is not (null or "rect" or "round")) throw new InvalidDataException("Invalid stroke-dash style.");
                int count = ReadCount("dots1", null), second = ReadCount("dots2", 0);
                return new OdfStrokeDashPattern(count, ReadLength("dots1-length"), ReadLength("distance"), second,
                    second > 0 ? ReadLength("dots2-length") : (OdfLength?)null, style == "round");
            } catch (ArgumentException exception) { throw new InvalidDataException("Unsupported stroke-dash definition '" + Name + "'.", exception); }
        }
        set {
            if (value == null) throw new ArgumentNullException(nameof(value));
            WritePattern(_element, value); _document.MarkPartDirty("styles.xml");
        }
    }
    internal static void WritePattern(XElement element, OdfStrokeDashPattern pattern) {
        element.SetAttributeValue(OdfNamespaces.Draw + "style", pattern.RoundCaps ? "round" : "rect");
        element.SetAttributeValue(OdfNamespaces.Draw + "dots1", pattern.DashCount.ToString(CultureInfo.InvariantCulture));
        element.SetAttributeValue(OdfNamespaces.Draw + "dots1-length", pattern.DashLength.ToString());
        element.SetAttributeValue(OdfNamespaces.Draw + "distance", pattern.Distance.ToString());
        element.SetAttributeValue(OdfNamespaces.Draw + "dots2", pattern.SecondDashCount > 0 ? pattern.SecondDashCount.ToString(CultureInfo.InvariantCulture) : null);
        element.SetAttributeValue(OdfNamespaces.Draw + "dots2-length", pattern.SecondDashCount > 0 ? pattern.SecondDashLength?.ToString() : null);
    }
    private int ReadCount(string name, int? fallback) {
        string? value = (string?)_element.Attribute(OdfNamespaces.Draw + name);
        if (value == null && fallback.HasValue) return fallback.Value;
        if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count)) throw new InvalidDataException("Missing or invalid dash count.");
        return count;
    }
    private OdfLength ReadLength(string name) {
        string? value = (string?)_element.Attribute(OdfNamespaces.Draw + name);
        if (string.IsNullOrWhiteSpace(value)) throw new NotSupportedException("Stroke-dash '" + Name + "' omits " + name + ".");
        return OdfLength.Parse(value!);
    }
}

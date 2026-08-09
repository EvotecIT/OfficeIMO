namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed common or automatic ODF style.</summary>
public sealed class OdfStyle {
    private readonly OdfDocument _document;
    private readonly XElement _element;

    internal OdfStyle(OdfDocument document, XElement element, string partPath, bool isAutomatic) {
        _document = document;
        _element = element;
        PartPath = partPath;
        IsAutomatic = isAutomatic;
    }

    /// <summary>Style name.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Style + "name") ?? string.Empty;
    /// <summary>Style family.</summary>
    public OdfStyleFamily Family {
        get {
            if (!OdfStyleRepository.TryParseFamily((string?)_element.Attribute(OdfNamespaces.Style + "family"), out OdfStyleFamily family)) {
                throw new InvalidDataException($"Style '{Name}' has an unsupported family.");
            }
            return family;
        }
    }
    /// <summary>Optional parent style name.</summary>
    public string? ParentStyleName {
        get => (string?)_element.Attribute(OdfNamespaces.Style + "parent-style-name");
        set => SetAttribute(_element, OdfNamespaces.Style + "parent-style-name", value);
    }
    /// <summary>True for an automatic style.</summary>
    public bool IsAutomatic { get; }
    /// <summary>Referenced number, currency, percentage, date, or time data style.</summary>
    public string? DataStyleName {
        get => (string?)_element.Attribute(OdfNamespaces.Style + "data-style-name");
        set => SetAttribute(_element, OdfNamespaces.Style + "data-style-name", value);
    }
    /// <summary>True when bold text is explicitly enabled by this style.</summary>
    public bool? Bold {
        get => ReadToggle(TextProperties, OdfNamespaces.Fo + "font-weight", "bold", "normal");
        set => WriteToggle(GetProperties(OdfNamespaces.Style + "text-properties"), OdfNamespaces.Fo + "font-weight", value, "bold", "normal");
    }
    /// <summary>True when italic text is explicitly enabled by this style.</summary>
    public bool? Italic {
        get => ReadToggle(TextProperties, OdfNamespaces.Fo + "font-style", "italic", "normal");
        set => WriteToggle(GetProperties(OdfNamespaces.Style + "text-properties"), OdfNamespaces.Fo + "font-style", value, "italic", "normal");
    }
    /// <summary>True when text underline is explicitly enabled by this style.</summary>
    public bool? Underline {
        get => ReadDecorationToggle(TextProperties, OdfNamespaces.Style + "text-underline-style");
        set => WriteDecorationToggle(OdfNamespaces.Style + "text-underline-style", value);
    }
    /// <summary>Whether this style explicitly uses an underline variant other than <c>solid</c> or <c>none</c>.</summary>
    public bool? UsesNonSolidUnderlineStyle => ReadNonSolidDecoration(
        TextProperties, OdfNamespaces.Style + "text-underline-style");
    /// <summary>True when text strike-through is explicitly enabled by this style.</summary>
    public bool? StrikeThrough {
        get => ReadDecorationToggle(TextProperties, OdfNamespaces.Style + "text-line-through-style");
        set => WriteDecorationToggle(OdfNamespaces.Style + "text-line-through-style", value);
    }
    /// <summary>Whether this style explicitly uses a line-through variant other than <c>solid</c> or <c>none</c>.</summary>
    public bool? UsesNonSolidLineThroughStyle => ReadNonSolidDecoration(
        TextProperties, OdfNamespaces.Style + "text-line-through-style");
    /// <summary>Explicit font size.</summary>
    public OdfLength? FontSize {
        get => ReadLength(TextProperties, OdfNamespaces.Fo + "font-size");
        set => SetAttribute(GetProperties(OdfNamespaces.Style + "text-properties"), OdfNamespaces.Fo + "font-size", value?.ToString());
    }
    /// <summary>Explicit font family.</summary>
    public string? FontFamily {
        get => (string?)TextProperties?.Attribute(OdfNamespaces.Fo + "font-family");
        set => SetAttribute(GetProperties(OdfNamespaces.Style + "text-properties"), OdfNamespaces.Fo + "font-family", value);
    }
    /// <summary>Explicit text color.</summary>
    public OdfColor? Color {
        get {
            string? value = (string?)TextProperties?.Attribute(OdfNamespaces.Fo + "color");
            return value == null ? (OdfColor?)null : OdfColor.Parse(value);
        }
        set => SetAttribute(GetProperties(OdfNamespaces.Style + "text-properties"), OdfNamespaces.Fo + "color", value?.ToString());
    }
    /// <summary>Explicit text background color, commonly used for highlighting.</summary>
    public OdfColor? TextBackgroundColor {
        get {
            TryGetTextBackgroundColor(out OdfColor? color);
            return color;
        }
        set => SetAttribute(GetProperties(OdfNamespaces.Style + "text-properties"), OdfNamespaces.Fo + "background-color", value?.ToString());
    }
    /// <summary>Explicit cell or paragraph background color.</summary>
    public OdfColor? BackgroundColor {
        get {
            TryGetBackgroundColor(out OdfColor? color);
            return color;
        }
        set {
            XName properties = Family == OdfStyleFamily.TableCell
                ? OdfNamespaces.Style + "table-cell-properties"
                : OdfNamespaces.Style + "paragraph-properties";
            SetAttribute(GetProperties(properties), OdfNamespaces.Fo + "background-color", value?.ToString());
        }
    }
    /// <summary>Explicit paragraph break-before value.</summary>
    public string? BreakBefore {
        get => (string?)ParagraphProperties?.Attribute(OdfNamespaces.Fo + "break-before");
        set => SetAttribute(GetProperties(OdfNamespaces.Style + "paragraph-properties"), OdfNamespaces.Fo + "break-before", value);
    }
    /// <summary>Explicit horizontal paragraph alignment.</summary>
    public string? TextAlign {
        get => (string?)ParagraphProperties?.Attribute(OdfNamespaces.Fo + "text-align");
        set => SetAttribute(GetProperties(OdfNamespaces.Style + "paragraph-properties"), OdfNamespaces.Fo + "text-align", value);
    }
    /// <summary>Explicit paragraph start margin.</summary>
    public OdfLength? MarginLeft {
        get => ReadLength(ParagraphProperties, OdfNamespaces.Fo + "margin-left");
        set => SetAttribute(GetProperties(OdfNamespaces.Style + "paragraph-properties"), OdfNamespaces.Fo + "margin-left", value?.ToString());
    }
    /// <summary>Explicit paragraph end margin.</summary>
    public OdfLength? MarginRight {
        get => ReadLength(ParagraphProperties, OdfNamespaces.Fo + "margin-right");
        set => SetAttribute(GetProperties(OdfNamespaces.Style + "paragraph-properties"), OdfNamespaces.Fo + "margin-right", value?.ToString());
    }
    /// <summary>Explicit paragraph top margin.</summary>
    public OdfLength? MarginTop {
        get => ReadLength(ParagraphProperties, OdfNamespaces.Fo + "margin-top");
        set => SetAttribute(GetProperties(OdfNamespaces.Style + "paragraph-properties"), OdfNamespaces.Fo + "margin-top", value?.ToString());
    }
    /// <summary>Explicit paragraph bottom margin.</summary>
    public OdfLength? MarginBottom {
        get => ReadLength(ParagraphProperties, OdfNamespaces.Fo + "margin-bottom");
        set => SetAttribute(GetProperties(OdfNamespaces.Style + "paragraph-properties"), OdfNamespaces.Fo + "margin-bottom", value?.ToString());
    }
    /// <summary>Explicit first-line paragraph indentation.</summary>
    public OdfLength? TextIndent {
        get => ReadLength(ParagraphProperties, OdfNamespaces.Fo + "text-indent");
        set => SetAttribute(GetProperties(OdfNamespaces.Style + "paragraph-properties"), OdfNamespaces.Fo + "text-indent", value?.ToString());
    }

    internal string PartPath { get; }
    internal XElement Element => _element;
    internal XElement? TextProperties => _element.Element(OdfNamespaces.Style + "text-properties");
    internal XElement? ParagraphProperties => _element.Element(OdfNamespaces.Style + "paragraph-properties");

    internal bool TryGetTextBackgroundColor(out OdfColor? color) => TryReadColorOverride(
        TextProperties, OdfNamespaces.Fo + "background-color", out color);

    internal bool TryGetBackgroundColor(out OdfColor? color) => TryReadColorOverride(
        _element.Element(OdfNamespaces.Style + "table-cell-properties") ?? ParagraphProperties,
        OdfNamespaces.Fo + "background-color", out color);

    internal XElement GetProperties(XName name) {
        XElement? properties = _element.Element(name);
        if (properties == null) {
            properties = new XElement(name);
            _element.Add(properties);
            _document.MarkPartDirty(PartPath);
        }
        return properties;
    }

    internal void SetProperty(XName propertyElement, XName attribute, string? value) {
        SetAttribute(GetProperties(propertyElement), attribute, value);
    }

    private void SetAttribute(XElement owner, XName name, object? value) {
        owner.SetAttributeValue(name, value);
        _document.MarkPartDirty(PartPath);
    }

    private static bool? ReadToggle(XElement? element, XName attribute, string trueValue, string falseValue) {
        string? value = (string?)element?.Attribute(attribute);
        if (string.Equals(value, trueValue, StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(value, falseValue, StringComparison.OrdinalIgnoreCase)) return false;
        return null;
    }

    private void WriteToggle(XElement element, XName attribute, bool? value, string trueValue, string falseValue) {
        SetAttribute(element, attribute, value.HasValue ? (value.Value ? trueValue : falseValue) : null);
    }

    private static bool? ReadDecorationToggle(XElement? element, XName attribute) {
        string? value = (string?)element?.Attribute(attribute);
        if (value == null) return null;
        return string.Equals(value, "none", StringComparison.OrdinalIgnoreCase) ? false : true;
    }

    private static bool? ReadNonSolidDecoration(XElement? element, XName attribute) {
        string? value = (string?)element?.Attribute(attribute);
        if (value == null) return null;
        return !string.Equals(value, "none", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(value, "solid", StringComparison.OrdinalIgnoreCase);
    }

    private void WriteDecorationToggle(XName attribute, bool? value) {
        SetAttribute(GetProperties(OdfNamespaces.Style + "text-properties"), attribute,
            value.HasValue ? (value.Value ? "solid" : "none") : null);
    }

    private static OdfLength? ReadLength(XElement? element, XName attribute) {
        string? value = (string?)element?.Attribute(attribute);
        return value == null ? (OdfLength?)null : OdfLength.Parse(value);
    }

    private static bool TryReadColorOverride(XElement? element, XName attribute, out OdfColor? color) {
        string? value = (string?)element?.Attribute(attribute);
        if (value == null) {
            color = null;
            return false;
        }
        color = string.Equals(value, "transparent", StringComparison.OrdinalIgnoreCase)
            ? (OdfColor?)null
            : OdfColor.Parse(value);
        return true;
    }
}

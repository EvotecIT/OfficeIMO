namespace OfficeIMO.OpenDocument;

// One operation owns the timestamp and parsed styles; no persistent cache can outlive native XML edits.
internal sealed class OdfDateTimeFieldProjection {
    private readonly OdfDocument _document;
    private readonly OdfDateTimeFieldProjectionOptions _options;
    private readonly Dictionary<(string Part, string Name), OdfDateTimeDataStyle> _styles = new();
    private readonly Dictionary<(string Part, string Name), string> _unsupportedStyles = new();
    internal OdfDateTimeFieldProjection(OdfDocument document, OdfDateTimeFieldProjectionOptions? options) {
        _document = document; _options = (options ?? new OdfDateTimeFieldProjectionOptions()).Snapshot();
    }
    internal bool EvaluatesValues => _options.Mode != OdfDateTimeFieldProjectionMode.CachedText;
    internal string Resolve(XElement element, OdfTextFieldKind kind, bool isFixed) {
        string valueName = kind == OdfTextFieldKind.Date ? "date-value" : "time-value";
        string adjustName = kind == OdfTextFieldKind.Date ? "date-adjust" : "time-adjust";
        foreach (XAttribute attribute in element.Attributes()) {
            if (attribute.IsNamespaceDeclaration || attribute.Name == XNamespace.Xml + "id" || attribute.Name == OdfNamespaces.Style + "data-style-name" ||
                attribute.Name.Namespace == OdfNamespaces.Text && (attribute.Name.LocalName == "fixed" || attribute.Name.LocalName == valueName || attribute.Name.LocalName == adjustName)) continue;
            throw new NotSupportedException("The date/time field has unsupported native attributes.");
        }
        OdfDateTimeFieldValue value;
        if (_options.Mode == OdfDateTimeFieldProjectionMode.RefreshDynamic && !isFixed) value = OdfDateTimeFieldValue.FromTimestamp(_options.RefreshTimestamp!.Value);
        else if (!OdfDateTimeFieldValue.TryParse((string?)element.Attribute(OdfNamespaces.Text + valueName), kind, out value))
            throw new NotSupportedException("The field has no saved value within the Gregorian date/time lexical profile.");
        if (!value.TryAdjust((string?)element.Attribute(OdfNamespaces.Text + adjustName), kind, out value))
            throw new NotSupportedException("The field adjustment is outside the bounded calendar/clock duration profile.");
        string? name = (string?)element.Attribute(OdfNamespaces.Style + "data-style-name");
        if (string.IsNullOrEmpty(name)) throw new NotSupportedException("An explicit data-style binding is required for date/time evaluation.");
        var key = (_document.GetPartPath(element), name!);
        if (_unsupportedStyles.TryGetValue(key, out string? reason)) throw new NotSupportedException(reason);
        if (!_styles.TryGetValue(key, out OdfDateTimeDataStyle? style)) {
            try {
                OdfDataStyle? definition = _document.Styles.FindDataStyle(name!, key.Item1);
                if (definition == null) throw new NotSupportedException("The bound date/time data style is missing.");
                style = OdfDateTimeDataStyle.Parse(definition, _options.DefaultCultureName); _styles.Add(key, style);
            } catch (Exception exception) when (exception is NotSupportedException or InvalidDataException) {
                _unsupportedStyles.Add(key, exception.Message); throw;
            }
        }
        try { return style.Format(value); }
        catch (ArgumentOutOfRangeException) { throw new NotSupportedException("Rounded date/time falls outside Gregorian years 1–9999."); }
    }
}

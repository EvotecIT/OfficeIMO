namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODS cell produced without expanding surrounding repeat runs.</summary>
public sealed class OdsCell {
    private readonly OdsDocument _document;
    private XElement _element;

    internal OdsCell(OdsDocument document, XElement element) { _document = document; _element = element; }

    /// <summary>True when this position is covered by a merged cell.</summary>
    public bool IsCovered => _element.Name == OdfNamespaces.Table + "covered-table-cell";
    /// <summary>Typed cached value.</summary>
    public OdsCellValue Value => ReadValue(_element);
    /// <summary>Decoded display text.</summary>
    public string Text => ReadDisplayText(_element);
    /// <summary>Raw OpenFormula expression, including any namespace prefix.</summary>
    public string? Formula {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "formula");
        set { EnsureEditable(); _element.SetAttributeValue(OdfNamespaces.Table + "formula", value); Dirty(); }
    }
    /// <summary>Referenced cell style name.</summary>
    public string? StyleName {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "style-name");
        set { EnsureEditable(); _element.SetAttributeValue(OdfNamespaces.Table + "style-name", value); Dirty(); }
    }
    /// <summary>Referenced number/date/time data style through the cell style.</summary>
    public string? NumberFormatName {
        get {
            OdfStyle? style = StyleName == null ? null : _document.Styles.Find(OdfStyleFamily.TableCell, StyleName);
            return style?.DataStyleName;
        }
        set { EnsureStyle().DataStyleName = value; }
    }
    /// <summary>Referenced content validation rule.</summary>
    public string? ValidationName {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "content-validation-name");
        set { EnsureEditable(); _element.SetAttributeValue(OdfNamespaces.Table + "content-validation-name", value); Dirty(); }
    }
    /// <summary>Explicit or inherited bold state.</summary>
    public bool? Bold {
        get => Resolve(style => style.Bold);
        set => EnsureStyle().Bold = value;
    }
    /// <summary>Explicit or inherited text color.</summary>
    public OdfColor? Color {
        get => Resolve(style => style.Color);
        set => EnsureStyle().Color = value;
    }

    /// <summary>Clears value content while retaining style, formula, validation, and annotations.</summary>
    public void ClearValue() {
        EnsureEditable();
        ClearValueAttributes();
        foreach (XElement paragraph in _element.Elements(OdfNamespaces.Text + "p").ToList()) paragraph.Remove();
        Dirty();
    }

    /// <summary>Sets a string value and display text.</summary>
    public void SetString(string? value) {
        EnsureEditable();
        string text = value ?? string.Empty;
        ClearValueAttributes();
        _element.SetAttributeValue(OdfNamespaces.Office + "value-type", "string");
        _element.SetAttributeValue(OdfNamespaces.Office + "string-value", text);
        ReplaceDisplayText(text);
        Dirty();
    }

    /// <summary>Sets a double value.</summary>
    public void SetNumber(double value) => SetNumeric("float", value.ToString("R", CultureInfo.InvariantCulture), value.ToString(CultureInfo.InvariantCulture));
    /// <summary>Sets a decimal without a binary floating-point round trip.</summary>
    public void SetDecimal(decimal value) => SetNumeric("float", value.ToString(CultureInfo.InvariantCulture), value.ToString(CultureInfo.InvariantCulture));

    /// <summary>Sets a boolean value.</summary>
    public void SetBoolean(bool value) {
        EnsureEditable(); ClearValueAttributes();
        string lexical = value ? "true" : "false";
        _element.SetAttributeValue(OdfNamespaces.Office + "value-type", "boolean");
        _element.SetAttributeValue(OdfNamespaces.Office + "boolean-value", lexical);
        ReplaceDisplayText(lexical.ToUpperInvariant()); Dirty();
    }

    /// <summary>Sets a local date or date-time value.</summary>
    public void SetDate(DateTime value) {
        string lexical = value.TimeOfDay == TimeSpan.Zero
            ? value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : value.ToString("yyyy-MM-ddTHH:mm:ss.fffffff", CultureInfo.InvariantCulture).TrimEnd('0').TrimEnd('.');
        SetDateLexical(lexical, value.ToString("d", CultureInfo.InvariantCulture));
    }

    /// <summary>Sets an offset-aware date-time value.</summary>
    public void SetDateTime(DateTimeOffset value) => SetDateLexical(value.ToString("o", CultureInfo.InvariantCulture), value.ToString("g", CultureInfo.InvariantCulture));

    /// <summary>Sets a time-of-day using the ODF duration representation.</summary>
    public void SetTime(TimeSpan value) {
        if (value < TimeSpan.Zero || value >= TimeSpan.FromDays(1)) throw new ArgumentOutOfRangeException(nameof(value), "Time of day must be within one day.");
        SetDurationCore(value);
    }

    /// <summary>Sets a duration. ODF stores both time-of-day and duration values as ISO durations.</summary>
    public void SetDuration(TimeSpan value) => SetDurationCore(value);

    /// <summary>Sets a percentage from a fractional value, where 0.25 represents 25%.</summary>
    public void SetPercentage(decimal value) => SetNumeric("percentage", value.ToString(CultureInfo.InvariantCulture),
        (value * 100m).ToString(CultureInfo.InvariantCulture) + "%");

    /// <summary>Sets a currency amount and ISO currency code.</summary>
    public void SetCurrency(decimal value, string currencyCode) {
        if (string.IsNullOrWhiteSpace(currencyCode)) throw new ArgumentException("Currency code cannot be empty.", nameof(currencyCode));
        EnsureEditable(); ClearValueAttributes();
        string lexical = value.ToString(CultureInfo.InvariantCulture);
        _element.SetAttributeValue(OdfNamespaces.Office + "value-type", "currency");
        _element.SetAttributeValue(OdfNamespaces.Office + "value", lexical);
        _element.SetAttributeValue(OdfNamespaces.Office + "currency", currencyCode.ToUpperInvariant());
        ReplaceDisplayText(lexical + " " + currencyCode.ToUpperInvariant()); Dirty();
    }

    /// <summary>Sets a hyperlink display value without fetching its target.</summary>
    public void SetHyperlink(string text, string href) {
        if (string.IsNullOrWhiteSpace(href)) throw new ArgumentException("Hyperlink target cannot be empty.", nameof(href));
        SetString(text);
        XElement paragraph = _element.Elements(OdfNamespaces.Text + "p").First();
        paragraph.RemoveNodes();
        var link = new XElement(OdfNamespaces.Text + "a",
            new XAttribute(OdfNamespaces.XLink + "type", "simple"),
            new XAttribute(OdfNamespaces.XLink + "href", href));
        OdfTextCodec.Append(link, text);
        paragraph.Add(link); Dirty();
    }

    internal static OdsCellValue ReadValue(XElement element) {
        string display = ReadDisplayText(element);
        string? type = (string?)element.Attribute(OdfNamespaces.Office + "value-type");
        switch (type) {
            case "string": return new OdsCellValue(OdsCellValueKind.String,
                (string?)element.Attribute(OdfNamespaces.Office + "string-value") ?? display, display);
            case "float": return new OdsCellValue(OdsCellValueKind.Number,
                (string?)element.Attribute(OdfNamespaces.Office + "value") ?? string.Empty, display);
            case "boolean": return new OdsCellValue(OdsCellValueKind.Boolean,
                (string?)element.Attribute(OdfNamespaces.Office + "boolean-value") ?? "false", display);
            case "date": return new OdsCellValue(OdsCellValueKind.Date,
                (string?)element.Attribute(OdfNamespaces.Office + "date-value") ?? string.Empty, display);
            case "time": return new OdsCellValue(OdsCellValueKind.Time,
                (string?)element.Attribute(OdfNamespaces.Office + "time-value") ?? string.Empty, display);
            case "percentage": return new OdsCellValue(OdsCellValueKind.Percentage,
                (string?)element.Attribute(OdfNamespaces.Office + "value") ?? string.Empty, display);
            case "currency": return new OdsCellValue(OdsCellValueKind.Currency,
                (string?)element.Attribute(OdfNamespaces.Office + "value") ?? string.Empty, display,
                (string?)element.Attribute(OdfNamespaces.Office + "currency"));
            default: return display.Length > 0 ? new OdsCellValue(OdsCellValueKind.String, display, display) : OdsCellValue.Empty;
        }
    }

    internal static bool IsEmpty(XElement element) => element.Name == OdfNamespaces.Table + "covered-table-cell" ||
        ((string?)element.Attribute(OdfNamespaces.Office + "value-type") == null &&
         (string?)element.Attribute(OdfNamespaces.Table + "formula") == null && ReadDisplayText(element).Length == 0);

    internal void SetSpans(long rows, long columns) {
        _element.SetAttributeValue(OdfNamespaces.Table + "number-rows-spanned", rows > 1 ? rows : (long?)null);
        _element.SetAttributeValue(OdfNamespaces.Table + "number-columns-spanned", columns > 1 ? columns : (long?)null);
        Dirty();
    }

    internal void ReplaceWithCoveredCell() {
        var covered = new XElement(OdfNamespaces.Table + "covered-table-cell");
        _element.ReplaceWith(covered); _element = covered; Dirty();
    }

    private static string ReadDisplayText(XElement element) => string.Join("\n", element.Elements(OdfNamespaces.Text + "p").Select(OdfTextCodec.Read));

    private void SetNumeric(string valueType, string lexical, string display) {
        EnsureEditable(); ClearValueAttributes();
        _element.SetAttributeValue(OdfNamespaces.Office + "value-type", valueType);
        _element.SetAttributeValue(OdfNamespaces.Office + "value", lexical);
        ReplaceDisplayText(display); Dirty();
    }

    private void SetDateLexical(string lexical, string display) {
        EnsureEditable(); ClearValueAttributes();
        _element.SetAttributeValue(OdfNamespaces.Office + "value-type", "date");
        _element.SetAttributeValue(OdfNamespaces.Office + "date-value", lexical);
        ReplaceDisplayText(display); Dirty();
    }

    private void SetDurationCore(TimeSpan value) {
        EnsureEditable(); ClearValueAttributes();
        string lexical = XmlConvert.ToString(value);
        _element.SetAttributeValue(OdfNamespaces.Office + "value-type", "time");
        _element.SetAttributeValue(OdfNamespaces.Office + "time-value", lexical);
        ReplaceDisplayText(value.ToString()); Dirty();
    }

    private void ReplaceDisplayText(string text) {
        foreach (XElement paragraph in _element.Elements(OdfNamespaces.Text + "p").ToList()) paragraph.Remove();
        var replacement = new XElement(OdfNamespaces.Text + "p");
        OdfTextCodec.Append(replacement, text);
        _element.Add(replacement);
    }

    private void ClearValueAttributes() {
        foreach (XName name in new[] {
            OdfNamespaces.Office + "value-type", OdfNamespaces.Office + "value", OdfNamespaces.Office + "string-value",
            OdfNamespaces.Office + "boolean-value", OdfNamespaces.Office + "date-value", OdfNamespaces.Office + "time-value",
            OdfNamespaces.Office + "currency"
        }) _element.Attribute(name)?.Remove();
    }

    private OdfStyle EnsureStyle() {
        EnsureEditable();
        return _document.Styles.EnsureAutomaticStyle(_element, OdfNamespaces.Table + "style-name", OdfStyleFamily.TableCell, "ofCell");
    }

    private T? Resolve<T>(Func<OdfStyle, T?> selector) where T : struct {
        OdfStyle? style = StyleName == null ? null : _document.Styles.Find(OdfStyleFamily.TableCell, StyleName);
        if (style == null) return null;
        foreach (OdfStyle candidate in _document.Styles.Resolve(style)) {
            T? value = selector(candidate); if (value.HasValue) return value;
        }
        return null;
    }

    private void EnsureEditable() {
        if (IsCovered) throw new InvalidOperationException("Covered cells cannot be edited directly. Edit the merged range anchor instead.");
    }
    private void Dirty() => _document.MarkPartDirty("content.xml");
}

/// <summary>A sparse repeated cell run.</summary>
public sealed class OdsCellRun {
    private readonly XElement _element;
    internal OdsCellRun(OdsDocument document, XElement element, long startColumn, long repeatCount) {
        _element = element; StartColumn = startColumn; RepeatCount = repeatCount;
    }
    /// <summary>Zero-based first logical column.</summary>
    public long StartColumn { get; }
    /// <summary>Number of logical cells represented by the run.</summary>
    public long RepeatCount { get; }
    /// <summary>True when this is a covered merged-cell run.</summary>
    public bool IsCovered => _element.Name == OdfNamespaces.Table + "covered-table-cell";
    /// <summary>Prototype value shared by the run.</summary>
    public OdsCellValue Value => OdsCell.ReadValue(_element);
    /// <summary>Prototype formula shared by the run.</summary>
    public string? Formula => (string?)_element.Attribute(OdfNamespaces.Table + "formula");
}

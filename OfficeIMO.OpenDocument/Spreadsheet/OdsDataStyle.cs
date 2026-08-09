namespace OfficeIMO.OpenDocument;

/// <summary>Basic spreadsheet number format kinds.</summary>
public enum OdsDataStyleKind {
    /// <summary>Decimal number.</summary>
    Number,
    /// <summary>Percentage.</summary>
    Percentage,
    /// <summary>Currency.</summary>
    Currency,
    /// <summary>Date.</summary>
    Date,
    /// <summary>Time.</summary>
    Time
}

/// <summary>An XML-backed ODS number, percentage, currency, date, or time style.</summary>
public sealed class OdsDataStyle {
    /// <summary>Maximum number of characters supported by an Excel custom number-format code.</summary>
    public const int MaximumExcelNumberFormatCodeLength = 255;

    internal OdsDataStyle(XElement element, OdsDataStyleKind kind) { Element = element; Kind = kind; }
    /// <summary>Style name referenced by a cell style.</summary>
    public string Name => (string?)Element.Attribute(OdfNamespaces.Style + "name") ?? string.Empty;
    /// <summary>Data style kind.</summary>
    public OdsDataStyleKind Kind { get; }
    /// <summary>Configured decimal places, or zero when the style has no decimal number component.</summary>
    public int DecimalPlaces => (int?)Element.Descendants(OdfNamespaces.Number + "number").FirstOrDefault()
        ?.Attribute(OdfNamespaces.Number + "decimal-places") ?? 0;
    /// <summary>Minimum integer digits requested by the number component.</summary>
    public int MinimumIntegerDigits => Math.Max(1,
        (int?)Element.Descendants(OdfNamespaces.Number + "number").FirstOrDefault()
            ?.Attribute(OdfNamespaces.Number + "min-integer-digits") ?? 1);
    /// <summary>Whether the style requests grouped thousands.</summary>
    public bool UsesGrouping => string.Equals(
        (string?)Element.Descendants(OdfNamespaces.Number + "number").FirstOrDefault()
            ?.Attribute(OdfNamespaces.Number + "grouping"),
        "true",
        StringComparison.OrdinalIgnoreCase);
    /// <summary>Visible currency symbol or code, when present.</summary>
    public string? CurrencySymbol => Element.Descendants(OdfNamespaces.Number + "currency-symbol")
        .Select(element => element.Value)
        .FirstOrDefault(value => value.Length > 0);

    /// <summary>Projects the represented common ODF style to an Excel number-format code.</summary>
    /// <exception cref="InvalidOperationException">The style cannot be represented within Excel's custom-format limit.</exception>
    public string ToExcelNumberFormatCode() {
        if (TryGetExcelNumberFormatCode(out string formatCode)) return formatCode;
        throw new InvalidOperationException(
            $"The ODF data style cannot be represented within Excel's {MaximumExcelNumberFormatCodeLength}-character custom number-format limit.");
    }

    /// <summary>Attempts to project the represented common ODF style to a bounded Excel number-format code.</summary>
    public bool TryGetExcelNumberFormatCode(out string formatCode) {
        formatCode = string.Empty;
        if (Kind == OdsDataStyleKind.Date || Kind == OdsDataStyleKind.Time) {
            return TryBuildDateTimeFormat(out formatCode);
        }
        if (!TryBuildNumericComponent(out string number)) return false;
        var builder = new System.Text.StringBuilder();
        bool wroteNumber = false;
        bool wrotePercentageScaling = false;
        foreach (XElement child in Element.Elements()) {
            if (child.Name == OdfNamespaces.Number + "number") {
                if (!TryAppend(builder, number)) return false;
                wroteNumber = true;
            } else if (child.Name == OdfNamespaces.Number + "currency-symbol") {
                if (!TryAppendExcelLiteral(builder, child.Value)) return false;
            } else if (child.Name == OdfNamespaces.Number + "text") {
                string text = child.Value;
                if (Kind == OdsDataStyleKind.Percentage && !wrotePercentageScaling) {
                    int percent = text.IndexOf('%');
                    if (percent >= 0) {
                        if (!TryAppendExcelLiteral(builder, text.Substring(0, percent)) ||
                            !TryAppend(builder, "%") ||
                            !TryAppendExcelLiteral(builder, text.Substring(percent + 1))) return false;
                        wrotePercentageScaling = true;
                    } else {
                        if (!TryAppendExcelLiteral(builder, text)) return false;
                    }
                } else {
                    if (!TryAppendExcelLiteral(builder, text)) return false;
                }
            }
        }
        if (!wroteNumber && !TryAppend(builder, number)) return false;
        if (Kind == OdsDataStyleKind.Percentage && !wrotePercentageScaling && !TryAppend(builder, "%")) return false;
        formatCode = builder.ToString();
        return true;
    }

    internal XElement Element { get; }

    private bool TryBuildNumericComponent(out string number) {
        number = string.Empty;
        XElement? component = Element.Descendants(OdfNamespaces.Number + "number").FirstOrDefault();
        if (!TryReadNonNegativeInteger(component, "min-integer-digits", 1, out int minimumIntegerDigits) ||
            !TryReadNonNegativeInteger(component, "decimal-places", 0, out int decimalPlaces)) return false;
        minimumIntegerDigits = Math.Max(1, minimumIntegerDigits);
        int groupingLength = UsesGrouping ? 4 : 0;
        int decimalSeparatorLength = decimalPlaces > 0 ? 1 : 0;
        if (minimumIntegerDigits > MaximumExcelNumberFormatCodeLength - groupingLength - decimalSeparatorLength ||
            decimalPlaces > MaximumExcelNumberFormatCodeLength - groupingLength - decimalSeparatorLength - minimumIntegerDigits) return false;
        number = (UsesGrouping ? "#,##" : string.Empty) + new string('0', minimumIntegerDigits);
        if (decimalPlaces > 0) number += "." + new string('0', decimalPlaces);
        return true;
    }

    private bool TryBuildDateTimeFormat(out string formatCode) {
        formatCode = string.Empty;
        var builder = new System.Text.StringBuilder();
        foreach (XElement child in Element.Elements()) {
            string style = (string?)child.Attribute(OdfNamespaces.Number + "style") ?? "short";
            if (child.Name == OdfNamespaces.Number + "year") {
                if (!TryAppend(builder, style == "long" ? "yyyy" : "yy")) return false;
            } else if (child.Name == OdfNamespaces.Number + "month") {
                bool textual = string.Equals(
                    (string?)child.Attribute(OdfNamespaces.Number + "textual"),
                    "true",
                    StringComparison.OrdinalIgnoreCase);
                if (!TryAppend(builder, textual
                    ? (style == "long" ? "mmmm" : "mmm")
                    : (style == "long" ? "mm" : "m"))) return false;
            } else if (child.Name == OdfNamespaces.Number + "day") {
                if (!TryAppend(builder, style == "long" ? "dd" : "d")) return false;
            } else if (child.Name == OdfNamespaces.Number + "hours") {
                if (!TryAppend(builder, style == "long" ? "hh" : "h")) return false;
            } else if (child.Name == OdfNamespaces.Number + "minutes") {
                if (!TryAppend(builder, style == "long" ? "mm" : "m")) return false;
            } else if (child.Name == OdfNamespaces.Number + "seconds") {
                if (!TryAppend(builder, style == "long" ? "ss" : "s") ||
                    !TryReadNonNegativeInteger(child, "decimal-places", 0, out int decimalPlaces)) return false;
                if (decimalPlaces > 0) {
                    if (decimalPlaces > MaximumExcelNumberFormatCodeLength - builder.Length - 1 ||
                        !TryAppend(builder, ".") || !TryAppend(builder, new string('0', decimalPlaces))) return false;
                }
            } else if (child.Name == OdfNamespaces.Number + "am-pm") {
                if (!TryAppend(builder, "AM/PM")) return false;
            } else if (child.Name == OdfNamespaces.Number + "text" && !TryAppendExcelLiteral(builder, child.Value)) return false;
        }
        formatCode = builder.Length == 0 ? (Kind == OdsDataStyleKind.Date ? "yyyy-mm-dd" : "hh:mm:ss") : builder.ToString();
        return true;
    }

    private static bool TryReadNonNegativeInteger(XElement? element, string localName, int fallback, out int value) {
        string? lexical = (string?)element?.Attribute(OdfNamespaces.Number + localName);
        if (lexical == null) {
            value = fallback;
            return true;
        }
        return int.TryParse(lexical, System.Globalization.NumberStyles.None,
            System.Globalization.CultureInfo.InvariantCulture, out value) && value >= 0;
    }

    private static bool TryAppendExcelLiteral(System.Text.StringBuilder builder, string value) {
        bool needsQuotes = !value.All(character => character == '-' || character == '/' || character == ':' || character == ' ' ||
                                                     character == '$' || character == '€' || character == '£' || character == '¥');
        if (!needsQuotes) return TryAppend(builder, value);
        long escapedLength = (long)value.Length + 2L + value.Count(character => character == '"');
        if (escapedLength > MaximumExcelNumberFormatCodeLength - builder.Length) return false;
        builder.Append('"').Append(value.Replace("\"", "\"\"")).Append('"');
        return true;
    }

    private static bool TryAppend(System.Text.StringBuilder builder, string value) {
        if (value.Length > MaximumExcelNumberFormatCodeLength - builder.Length) return false;
        builder.Append(value);
        return true;
    }
}

public sealed partial class OdsDocument {
    /// <summary>Number and date/time styles available in this spreadsheet.</summary>
    public IReadOnlyList<OdsDataStyle> DataStyles {
        get {
            var result = new List<OdsDataStyle>();
            foreach (string partPath in new[] { "content.xml", "styles.xml" }) {
                if (!Package.ContainsEntry(partPath)) continue;
                XElement? root = GetXml(partPath).Root;
                foreach (XElement container in root?.Elements().Where(element =>
                             element.Name == OdfNamespaces.Office + "automatic-styles" || element.Name == OdfNamespaces.Office + "styles") ?? Enumerable.Empty<XElement>()) {
                    foreach (XElement element in container.Elements()) {
                        if (TryGetKind(element.Name, out OdsDataStyleKind kind)) result.Add(new OdsDataStyle(element, kind));
                    }
                }
            }
            return result;
        }
    }

    /// <summary>Adds a decimal number style.</summary>
    public OdsDataStyle AddNumberStyle(string name, int decimalPlaces = 2, bool useGrouping = false) {
        if (decimalPlaces < 0) throw new ArgumentOutOfRangeException(nameof(decimalPlaces));
        return AddDataStyle(name, OdsDataStyleKind.Number,
            new XElement(OdfNamespaces.Number + "number",
                new XAttribute(OdfNamespaces.Number + "decimal-places", decimalPlaces),
                new XAttribute(OdfNamespaces.Number + "grouping", useGrouping),
                new XAttribute(OdfNamespaces.Number + "min-integer-digits", 1)));
    }

    /// <summary>Adds a percentage style.</summary>
    public OdsDataStyle AddPercentageStyle(string name, int decimalPlaces = 2, bool useGrouping = false) {
        if (decimalPlaces < 0) throw new ArgumentOutOfRangeException(nameof(decimalPlaces));
        return AddDataStyle(name, OdsDataStyleKind.Percentage,
            new XElement(OdfNamespaces.Number + "number",
                new XAttribute(OdfNamespaces.Number + "decimal-places", decimalPlaces),
                new XAttribute(OdfNamespaces.Number + "grouping", useGrouping),
                new XAttribute(OdfNamespaces.Number + "min-integer-digits", 1)),
            new XElement(OdfNamespaces.Number + "text", "%"));
    }

    /// <summary>Adds a currency style with a visible currency symbol or code.</summary>
    public OdsDataStyle AddCurrencyStyle(string name, string currencySymbol, int decimalPlaces = 2, bool useGrouping = false) {
        if (string.IsNullOrWhiteSpace(currencySymbol)) throw new ArgumentException("Currency symbol cannot be empty.", nameof(currencySymbol));
        if (decimalPlaces < 0) throw new ArgumentOutOfRangeException(nameof(decimalPlaces));
        return AddDataStyle(name, OdsDataStyleKind.Currency,
            new XElement(OdfNamespaces.Number + "currency-symbol", currencySymbol),
            new XElement(OdfNamespaces.Number + "text", " "),
            new XElement(OdfNamespaces.Number + "number",
                new XAttribute(OdfNamespaces.Number + "decimal-places", decimalPlaces),
                new XAttribute(OdfNamespaces.Number + "grouping", useGrouping),
                new XAttribute(OdfNamespaces.Number + "min-integer-digits", 1)));
    }

    /// <summary>Adds an ISO-style year-month-day date format.</summary>
    public OdsDataStyle AddDateStyle(string name) => AddDataStyle(name, OdsDataStyleKind.Date,
        new XElement(OdfNamespaces.Number + "year", new XAttribute(OdfNamespaces.Number + "style", "long")),
        new XElement(OdfNamespaces.Number + "text", "-"),
        new XElement(OdfNamespaces.Number + "month", new XAttribute(OdfNamespaces.Number + "style", "long")),
        new XElement(OdfNamespaces.Number + "text", "-"),
        new XElement(OdfNamespaces.Number + "day", new XAttribute(OdfNamespaces.Number + "style", "long")));

    /// <summary>Adds a 24-hour time format with seconds.</summary>
    public OdsDataStyle AddTimeStyle(string name) => AddDataStyle(name, OdsDataStyleKind.Time,
        new XElement(OdfNamespaces.Number + "hours", new XAttribute(OdfNamespaces.Number + "style", "long")),
        new XElement(OdfNamespaces.Number + "text", ":"),
        new XElement(OdfNamespaces.Number + "minutes", new XAttribute(OdfNamespaces.Number + "style", "long")),
        new XElement(OdfNamespaces.Number + "text", ":"),
        new XElement(OdfNamespaces.Number + "seconds", new XAttribute(OdfNamespaces.Number + "style", "long")));

    private OdsDataStyle AddDataStyle(string name, OdsDataStyleKind kind, params XElement[] children) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Data style name cannot be empty.", nameof(name));
        if (DataStyles.Any(style => string.Equals(style.Name, name, StringComparison.Ordinal))) {
            throw new InvalidOperationException($"A data style named '{name}' already exists.");
        }
        XElement root = GetXml("content.xml").Root ?? throw new InvalidDataException("OpenDocument content has no root element.");
        XElement container = root.Element(OdfNamespaces.Office + "automatic-styles") ?? throw new InvalidDataException("OpenDocument content has no automatic styles.");
        XName elementName;
        switch (kind) {
            case OdsDataStyleKind.Number: elementName = OdfNamespaces.Number + "number-style"; break;
            case OdsDataStyleKind.Percentage: elementName = OdfNamespaces.Number + "percentage-style"; break;
            case OdsDataStyleKind.Currency: elementName = OdfNamespaces.Number + "currency-style"; break;
            case OdsDataStyleKind.Date: elementName = OdfNamespaces.Number + "date-style"; break;
            default: elementName = OdfNamespaces.Number + "time-style"; break;
        }
        var element = new XElement(elementName, new XAttribute(OdfNamespaces.Style + "name", name), children);
        container.Add(element); MarkPartDirty("content.xml");
        return new OdsDataStyle(element, kind);
    }

    private static bool TryGetKind(XName name, out OdsDataStyleKind kind) {
        if (name == OdfNamespaces.Number + "number-style") kind = OdsDataStyleKind.Number;
        else if (name == OdfNamespaces.Number + "percentage-style") kind = OdsDataStyleKind.Percentage;
        else if (name == OdfNamespaces.Number + "currency-style") kind = OdsDataStyleKind.Currency;
        else if (name == OdfNamespaces.Number + "date-style") kind = OdsDataStyleKind.Date;
        else if (name == OdfNamespaces.Number + "time-style") kind = OdsDataStyleKind.Time;
        else { kind = default; return false; }
        return true;
    }
}

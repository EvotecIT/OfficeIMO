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
    internal OdsDataStyle(XElement element, OdsDataStyleKind kind) { Element = element; Kind = kind; }
    /// <summary>Style name referenced by a cell style.</summary>
    public string Name => (string?)Element.Attribute(OdfNamespaces.Style + "name") ?? string.Empty;
    /// <summary>Data style kind.</summary>
    public OdsDataStyleKind Kind { get; }
    internal XElement Element { get; }
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
    public OdsDataStyle AddNumberStyle(string name, int decimalPlaces = 2) {
        if (decimalPlaces < 0) throw new ArgumentOutOfRangeException(nameof(decimalPlaces));
        return AddDataStyle(name, OdsDataStyleKind.Number,
            new XElement(OdfNamespaces.Number + "number",
                new XAttribute(OdfNamespaces.Number + "decimal-places", decimalPlaces),
                new XAttribute(OdfNamespaces.Number + "min-integer-digits", 1)));
    }

    /// <summary>Adds a percentage style.</summary>
    public OdsDataStyle AddPercentageStyle(string name, int decimalPlaces = 2) {
        if (decimalPlaces < 0) throw new ArgumentOutOfRangeException(nameof(decimalPlaces));
        return AddDataStyle(name, OdsDataStyleKind.Percentage,
            new XElement(OdfNamespaces.Number + "number",
                new XAttribute(OdfNamespaces.Number + "decimal-places", decimalPlaces),
                new XAttribute(OdfNamespaces.Number + "min-integer-digits", 1)),
            new XElement(OdfNamespaces.Number + "text", "%"));
    }

    /// <summary>Adds a currency style with a visible currency symbol or code.</summary>
    public OdsDataStyle AddCurrencyStyle(string name, string currencySymbol, int decimalPlaces = 2) {
        if (string.IsNullOrWhiteSpace(currencySymbol)) throw new ArgumentException("Currency symbol cannot be empty.", nameof(currencySymbol));
        if (decimalPlaces < 0) throw new ArgumentOutOfRangeException(nameof(decimalPlaces));
        return AddDataStyle(name, OdsDataStyleKind.Currency,
            new XElement(OdfNamespaces.Number + "currency-symbol", currencySymbol),
            new XElement(OdfNamespaces.Number + "text", " "),
            new XElement(OdfNamespaces.Number + "number",
                new XAttribute(OdfNamespaces.Number + "decimal-places", decimalPlaces),
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

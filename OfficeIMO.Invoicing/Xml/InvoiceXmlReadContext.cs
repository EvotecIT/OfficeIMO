using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

/// <summary>Tracks every consumed XML element and attribute so unsupported input cannot disappear unnoticed.</summary>
internal sealed class InvoiceXmlReadContext {
    private readonly HashSet<XObject> _consumed = new HashSet<XObject>();
    private readonly HashSet<XElement> _scalarValues = new HashSet<XElement>();
    private readonly List<InvoiceDiagnostic> _diagnostics = new List<InvoiceDiagnostic>();
    private int _modelItems;
    internal void AddTo<T>(ICollection<T> collection, T item) {
        if (++_modelItems > InvoiceModelLimits.MaximumCollectionItems)
            throw new InvalidDataException("Parsed invoice exceeds 50,000 collection items.");
        collection.Add(item);
    }
    internal XElement? Child(XElement? parent, XName name) {
        if (parent == null) return null;
        XElement? child = InvoiceXml.Unique(parent, name);
        if (child != null) _consumed.Add(child);
        return child;
    }
    internal IEnumerable<XElement> Children(XElement? parent, XName name) {
        if (parent == null) yield break;
        foreach (XElement child in parent.Elements(name)) { _consumed.Add(child); yield return child; }
    }
    internal void Consume(XObject node) => _consumed.Add(node);
    internal string? Attribute(XElement? element, XName name) {
        XAttribute? attribute = element?.Attribute(name);
        if (attribute == null) return null;
        _consumed.Add(attribute); return attribute.Value;
    }
    internal string? Value(XElement? element) {
        if (element == null) return null;
        _consumed.Add(element);
        if (element.HasElements) throw new InvalidDataException("Invoice scalar contains child elements: " + Path(element));
        _scalarValues.Add(element);
        return element.Value;
    }
    internal string? Text(XElement? parent, XName name) => Value(Child(parent, name));
    internal string Required(XElement? parent, XName name) => Text(parent, name) ?? string.Empty;
    internal XElement Require(XElement? parent, XName name) => Child(parent, name) ?? throw new InvalidDataException("Required invoice element is missing: " + name + ".");
    internal decimal RequiredDecimal(XElement? element) => Decimal(element) ?? throw new InvalidDataException("Required invoice decimal is missing.");
    internal decimal? Decimal(XElement? element) {
        string? value = Value(element);
        if (value == null) return null;
        if (!decimal.TryParse(value, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite, CultureInfo.InvariantCulture, out decimal result))
            throw new InvalidDataException("Invalid decimal at " + Path(element!) + ".");
        if (CanonicalDecimal(value) != CanonicalDecimal(result.ToString(CultureInfo.InvariantCulture)))
            throw new InvalidDataException("Decimal cannot be represented exactly at " + Path(element!) + ".");
        return result;
    }
    private static string CanonicalDecimal(string value) {
        string text = value.Trim();
        bool negative = text.StartsWith("-", StringComparison.Ordinal);
        if (text.StartsWith("+", StringComparison.Ordinal) || negative) text = text.Substring(1);
        int point = text.IndexOf('.');
        string whole = (point < 0 ? text : text.Substring(0, point)).TrimStart('0');
        string fraction = point < 0 ? string.Empty : text.Substring(point + 1).TrimEnd('0');
        if (whole.Length == 0) whole = "0";
        return (negative && (whole != "0" || fraction.Length != 0) ? "-" : string.Empty) + whole +
            (fraction.Length == 0 ? string.Empty : "." + fraction);
    }
    internal decimal? Decimal(XElement? parent, XName name) => Decimal(Child(parent, name));
    internal decimal? Money(XElement? parent, XName name, string? currency, bool requireCurrency = false) {
        XElement? element = Child(parent, name);
        string? actual = Attribute(element, "currencyID");
        if (element != null && (actual != null && actual != currency || requireCurrency && actual == null))
            Loss(element, "Amount currency is missing or differs from the invoice currency.");
        return Decimal(element);
    }
    internal decimal RequiredMoney(XElement? parent, XName name, string? currency, bool requireCurrency = false) =>
        Money(parent, name, currency, requireCurrency) ?? throw new InvalidDataException("Required invoice amount is missing: " + name + ".");
    internal bool Boolean(XElement? element) {
        string? value = Value(element);
        if (value == "true" || value == "1") return true;
        if (value == "false" || value == "0") return false;
        throw new InvalidDataException("Expected an XML boolean at " + (element == null ? "missing indicator" : Path(element)) + ".");
    }
    internal DateTime? Date(XElement? element, bool cii = false) {
        string? value = Value(element);
        if (value == null) return null;
        string format = cii ? "yyyyMMdd" : "yyyy-MM-dd";
        string? code = Attribute(element, "format");
        if (cii && code != "102") throw new InvalidDataException("Unsupported CII date format at " + Path(element!) + ".");
        if (!DateTime.TryParseExact(value, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
            throw new InvalidDataException("Invalid date at " + Path(element!) + ".");
        return result;
    }
    internal InvoiceIdentifier? Identifier(XElement? element, string schemeAttribute = "schemeID") {
        string? value = Value(element);
        return value == null ? null : new InvoiceIdentifier(value, Attribute(element, schemeAttribute));
    }
    internal void Expected(XElement? element, string expected) {
        string? actual = Value(element);
        if (actual != null && actual != expected) Loss(element!, "Unsupported value '" + actual + "'; expected '" + expected + "'.");
    }
    internal void Loss(XElement element, string message) => AddDiagnostic(new InvoiceDiagnostic("INV-UNMAPPED", message, Path(element)));
    private void AddDiagnostic(InvoiceDiagnostic diagnostic) {
        if (_diagnostics.Count >= 1000) throw new InvalidDataException("Invoice has more than 1,000 mapping diagnostics.");
        _diagnostics.Add(diagnostic);
    }
    internal IReadOnlyList<InvoiceDiagnostic> Finish(XElement root) {
        foreach (XElement element in root.DescendantsAndSelf()) {
            if (!_consumed.Contains(element) && (element.Parent == null || _consumed.Contains(element.Parent)))
                Loss(element, "Element is outside the supported semantic mapping.");
            if (_consumed.Contains(element)) {
                foreach (XAttribute attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration && !_consumed.Contains(a)))
                    AddDiagnostic(new InvoiceDiagnostic("INV-UNMAPPED", "Attribute is outside the supported semantic mapping.", Path(element) + "/@" + attribute.Name));
                if (!_scalarValues.Contains(element) && element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                    Loss(element, "Text in an invoice container is outside the supported mapping.");
            }
        }
        return _diagnostics.AsReadOnly();
    }
    private static string Path(XElement element) => "/" + string.Join("/", element.AncestorsAndSelf().Reverse().Select(e =>
        "{" + e.Name.NamespaceName + "}" + e.Name.LocalName + "[" + (e.ElementsBeforeSelf(e.Name).Count() + 1) + "]"));
}

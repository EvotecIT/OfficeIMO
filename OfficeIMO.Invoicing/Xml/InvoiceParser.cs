using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

/// <summary>Bounded native invoice parser with explicit unsupported-data reporting.</summary>
public static partial class InvoiceParser {
    private static readonly XNamespace Ram = InvoiceXml.Ram;
    private static readonly XNamespace Rsm = InvoiceXml.Rsm;
    private static readonly XNamespace Udt = "urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100";
    private static readonly XNamespace Qdt = "urn:un:unece:uncefact:data:standard:QualifiedDataType:100";
    private static readonly XNamespace Cbc = InvoiceXml.Cbc;
    private static readonly XNamespace Cac = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";

    /// <summary>Reads an editable model and reports every unmapped element or attribute. Parsing alone does not establish compliance.</summary>
    public static InvoiceReadResult Read(byte[] xml) {
        if (xml == null) throw new ArgumentNullException(nameof(xml));
        if (xml.Length == 0 || xml.Length > InvoiceProfileDeclaration.MaximumXmlBytes) throw new InvalidDataException("Invoice XML must contain between 1 byte and 16 MiB.");
        byte[] snapshot = (byte[])xml.Clone();
        XDocument document = InvoiceXml.Parse(snapshot);
        InvoiceProfileDeclaration declaration = InvoiceProfileDeclaration.Read(document);
        var context = new InvoiceXmlReadContext();
        XElement root = document.Root!;
        context.Consume(root);
        Invoice invoice = declaration.Syntax == InvoiceSyntax.Cii ? ReadCii(root, context) : ReadUbl(root, context);
        if (!declaration.Profile.HasValue) context.Loss(root, "The source guideline is not in the supported profile catalogue.");
        return new InvoiceReadResult(invoice, declaration, snapshot, context.Finish(root));
    }

    private static DateTime? CiiDate(InvoiceXmlReadContext c, XElement? parent, string name, bool qualified = false) =>
        c.Date(c.Child(c.Child(parent, Ram + name), (qualified ? Qdt : Udt) + "DateTimeString"), true);
    private static InvoicePeriod? CiiPeriod(InvoiceXmlReadContext c, XElement? parent) {
        XElement? element = c.Child(parent, Ram + "BillingSpecifiedPeriod");
        return element == null ? null : new InvoicePeriod { Start = CiiDate(c, element, "StartDateTime"), End = CiiDate(c, element, "EndDateTime") };
    }
    private static InvoicePeriod? UblPeriod(InvoiceXmlReadContext c, XElement? parent, out string? taxPointCode) {
        XElement? element = c.Child(parent, Cac + "InvoicePeriod");
        taxPointCode = c.Text(element, Cbc + "DescriptionCode");
        if (element == null) return null;
        DateTime? start = c.Date(c.Child(element, Cbc + "StartDate")), end = c.Date(c.Child(element, Cbc + "EndDate"));
        return start == null && end == null ? null : new InvoicePeriod { Start = start, End = end };
    }
    private static byte[]? Binary(InvoiceXmlReadContext c, XElement? element) {
        string? value = c.Value(element);
        if (value == null) return null;
        try {
            byte[] data = System.Convert.FromBase64String(value);
            if (data.Length > 8 * 1024 * 1024) throw new InvalidDataException("Supporting document exceeds 8 MiB.");
            return data;
        } catch (FormatException exception) { throw new InvalidDataException("Supporting document contains invalid Base64.", exception); }
    }
    private static void Agree(InvoiceXmlReadContext c, XElement source, string? expected, string? actual, string field) {
        if (expected != actual) c.Loss(source, "Multiple " + field + " values cannot be represented by one model field.");
    }
}

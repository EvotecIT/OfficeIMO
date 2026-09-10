using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

/// <summary>Bounded, namespace-aware XML operations shared by invoice readers and profile inspection.</summary>
internal static class InvoiceXml {
    internal static readonly XNamespace Rsm = "urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100";
    internal static readonly XNamespace Ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
    internal static readonly XNamespace Cbc = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";
    internal static readonly XNamespace UblInvoice = "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2";
    internal static readonly XNamespace UblCreditNote = "urn:oasis:names:specification:ubl:schema:xsd:CreditNote-2";

    internal static XDocument Parse(byte[] bytes) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (bytes.Length == 0 || bytes.Length > InvoiceProfileDeclaration.MaximumXmlBytes)
            throw new InvalidDataException("Invoice XML must contain between 1 byte and 16 MiB.");
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = InvoiceProfileDeclaration.MaximumXmlBytes,
            IgnoreWhitespace = false
        };
        // Validate depth before materializing the tree, including deeply nested unknown extensions.
        using (var stream = new MemoryStream(bytes, false))
        using (var reader = XmlReader.Create(stream, settings)) {
            while (reader.Read()) {
                if (reader.Depth > 128) throw new InvalidDataException("Invoice XML exceeds the maximum depth of 128.");
            }
        }
        using (var stream = new MemoryStream(bytes, false))
        using (var reader = XmlReader.Create(stream, settings))
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
    }

    internal static XElement? Unique(XElement parent, XName name) {
        XElement[] matches = parent.Elements(name).Take(2).ToArray();
        if (matches.Length > 1) throw new InvalidDataException("Ambiguous invoice field: " + name + ".");
        return matches.Length == 0 ? null : matches[0];
    }

    internal static string Scalar(XElement element) {
        if (element.HasElements) throw new InvalidDataException("Invoice scalar contains child elements: " + element.Name + ".");
        return element.Value.Trim();
    }
}

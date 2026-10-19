using System.Xml.Linq;

namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private const string PdfAIdNamespace = "http://www.aiim.org/pdfa/ns/id/";

    private static string? XmpIdentificationValue(string? xml, string namespaceUri, string name) {
        XElement? root = ReadIdentificationXml(xml);
        XName key = XName.Get(name, namespaceUri);
        return root?.DescendantsAndSelf().SelectMany(element =>
            element.Attributes(key).Select(attribute => attribute.Value)
                .Concat(element.Name == key ? new[] { element.Value } : Array.Empty<string>())).FirstOrDefault();
    }

    private static XElement? ReadIdentificationXml(string? xml) {
        if (string.IsNullOrWhiteSpace(xml)) return null;
        try {
            using var reader = System.Xml.XmlReader.Create(new StringReader(xml), new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null
            });
            return XElement.Load(reader);
        } catch (System.Xml.XmlException) { return null; }
    }

    private static void AddPdfIdentificationReadbackRequirements(List<PdfComplianceRequirement> requirements, PdfXmpMetadataInfo? xmp, int part) {
        if (part == 4) {
            Add(requirements, "readback-pdfa-revision", "PDF/A-4 revision",
                XmpIdentificationValue(xmp?.RawXml, PdfAIdNamespace, "rev") == "2020",
                "PDF/A-4 revision is 2020.", "PDF/A-4 requires pdfaid:rev=2020.");
        } else if (xmp?.PdfUaPart.HasValue == true) {
            Add(requirements, "readback-pdfua-extension-schema", "PDF/A extension schema for PDF/UA identification",
                HasPdfUaExtensionSchema(xmp.RawXml),
                "The PDF/UA properties are declared in the PDF/A extension schema.",
                "PDF/A-2 and PDF/A-3 require pdfaExtension:schemas declarations for every pdfuaid property.");
        }
    }

    private static bool HasPdfUaExtensionSchema(string? xml) {
        XElement? root = ReadIdentificationXml(xml);
        if (root == null) return false;
        XNamespace extension = "http://www.aiim.org/pdfa/ns/extension/";
        XNamespace schema = "http://www.aiim.org/pdfa/ns/schema#";
        XNamespace property = "http://www.aiim.org/pdfa/ns/property#";
        XNamespace rdf = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
        XNamespace ua = PdfUaIdentification.NamespaceUri;
        string[] used = root.DescendantsAndSelf().SelectMany(element =>
            element.Attributes().Where(a => a.Name.Namespace == ua).Select(a => a.Name.LocalName)
                .Concat(element.Name.Namespace == ua ? new[] { element.Name.LocalName } : Array.Empty<string>()))
            .Distinct(StringComparer.Ordinal).ToArray();
        foreach (XElement item in root.Descendants(extension + "schemas").Elements(rdf + "Bag").Elements(rdf + "li")) {
            if ((string?)item.Element(schema + "namespaceURI") != PdfUaIdentification.NamespaceUri ||
                string.IsNullOrWhiteSpace((string?)item.Element(schema + "schema")) ||
                string.IsNullOrWhiteSpace((string?)item.Element(schema + "prefix"))) continue;
            XElement[] definitions = item.Elements(schema + "property").Elements(rdf + "Seq").Elements(rdf + "li").ToArray();
            if (used.All(name => definitions.Any(d => (string?)d.Element(property + "name") == name &&
                (string?)d.Element(property + "valueType") == (name == "part" ? "Integer" : "Text") &&
                (string?)d.Element(property + "category") == "internal" &&
                !string.IsNullOrWhiteSpace((string?)d.Element(property + "description"))))) return true;
        }
        return false;
    }
}

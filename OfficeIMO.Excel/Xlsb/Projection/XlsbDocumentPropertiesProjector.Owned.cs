using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel.Xlsb.Projection;

internal static partial class XlsbDocumentPropertiesProjector {
    private static readonly XNamespace XmlSchemaInstance = "http://www.w3.org/2001/XMLSchema-instance";
    private static readonly XNamespace Core = CoreNamespace;
    private static readonly XNamespace Dc = DublinCoreNamespace;
    private static readonly XNamespace Terms = DublinCoreTermsNamespace;
    private static readonly XName[] CoreNames = {
        Dc + "title", Dc + "subject", Dc + "creator", Core + "keywords",
        Dc + "description", Core + "lastModifiedBy", Core + "revision",
        Core + "category", Core + "version", Core + "contentStatus",
        Dc + "identifier", Dc + "language", Core + "lastPrinted", Terms + "created", Terms + "modified"
    };

    /// <summary>Installs the supported native metadata without loading an Excel model.</summary>
    internal static void Apply(OfficeOpcEditablePackage package, XlsbWorkbook workbook) {
        if (workbook.CorePropertiesXml != null) {
            if (TryLoadRoot(workbook.CorePropertiesXml, CoreNamespace, "coreProperties", out XElement? core)) {
                InstallCore(package, core!);
            } else {
                Skip(workbook, "core");
            }
        }
        if (workbook.ExtendedPropertiesXml != null) {
            if (TryLoadRoot(workbook.ExtendedPropertiesXml, ExtendedNamespace, "Properties", out _)) {
                AddPropertyPart(package, "docProps/app.xml", XlsbDocumentPropertyParts.ExtendedRelationshipType,
                    XlsbDocumentPropertyParts.ExtendedContentType, workbook.ExtendedPropertiesXml);
            } else {
                Skip(workbook, "extended");
            }
        }
    }

    private static void InstallCore(OfficeOpcEditablePackage package, XElement source) {
        var values = new Dictionary<XName, XElement>();
        foreach (XElement child in source.Elements()) {
            if (Array.IndexOf(CoreNames, child.Name) < 0) continue;
            if (child.Name == Core + "lastPrinted" || child.Name.Namespace == Terms) {
                if (ParseDate(child.Value) is not DateTime date) continue;
                var normalized = new XElement(child.Name, date.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture));
                if (child.Name.Namespace == Terms) normalized.SetAttributeValue(XmlSchemaInstance + "type", "dcterms:W3CDTF");
                values[child.Name] = normalized;
            } else {
                values[child.Name] = new XElement(child.Name, child.Value);
            }
        }
        if (values.Count == 0) return;
        var root = new XElement(Core + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "cp", CoreNamespace),
            new XAttribute(XNamespace.Xmlns + "dc", DublinCoreNamespace),
            new XAttribute(XNamespace.Xmlns + "dcterms", DublinCoreTermsNamespace),
            new XAttribute(XNamespace.Xmlns + "xsi", XmlSchemaInstance.NamespaceName));
        foreach (XName name in CoreNames) {
            if (values.TryGetValue(name, out XElement? element)) root.Add(element);
        }

        using var bytes = new MemoryStream();
        using (XmlWriter writer = XmlWriter.Create(bytes, new XmlWriterSettings { Encoding = new UTF8Encoding(false) })) root.WriteTo(writer);
        AddPropertyPart(package, "docProps/core.xml", XlsbDocumentPropertyParts.CoreRelationshipType,
            XlsbDocumentPropertyParts.CoreContentType, bytes.ToArray());
    }

    private static void AddPropertyPart(OfficeOpcEditablePackage package, string name, string relationshipType,
        string contentType, byte[] bytes) {
        package.CreateBinaryPart("/" + name, contentType, bytes);
        package.AddRelationship(null, relationshipType, name);
    }
}

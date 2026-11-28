using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private const string StreamingCoreNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private const string StreamingDcNamespace = "http://purl.org/dc/elements/1.1/";
    private const string StreamingTermsNamespace = "http://purl.org/dc/terms/";
    private const string StreamingSchemaInstance = "http://www.w3.org/2001/XMLSchema-instance";
    private const string StreamingAppNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    private static (string Name, string Namespace) ResolveStreamingProperty(string name) => name.ToUpperInvariant() switch {
        "TITLE" => ("title", StreamingDcNamespace),
        "SUBJECT" => ("subject", StreamingDcNamespace),
        "AUTHOR" or "CREATOR" => ("creator", StreamingDcNamespace),
        "DESCRIPTION" => ("description", StreamingDcNamespace),
        "KEYWORDS" => ("keywords", StreamingCoreNamespace),
        "CATEGORY" => ("category", StreamingCoreNamespace),
        "REVISION" => ("revision", StreamingCoreNamespace),
        "LASTMODIFIEDBY" => ("lastModifiedBy", StreamingCoreNamespace),
        "VERSION" => ("version", StreamingCoreNamespace),
        "CREATED" => ("created", StreamingTermsNamespace),
        "MODIFIED" => ("modified", StreamingTermsNamespace),
        "LASTPRINTED" => ("lastPrinted", StreamingCoreNamespace),
        "APPLICATION" or "APPLICATIONNAME" => ("Application", StreamingAppNamespace),
        "COMPANY" => ("Company", StreamingAppNamespace),
        "MANAGER" => ("Manager", StreamingAppNamespace),
        _ => throw new ArgumentException($"'{name}' is not a supported Excel document property.", "properties"),
    };

    private static void ValidateStreamingDocumentProperties(IReadOnlyDictionary<string, object?>? properties) {
        if (properties == null) return;
        foreach (string name in properties.Keys) _ = ResolveStreamingProperty(name);
    }

    private static (string? Core, string? Application) WriteStreamingDocumentProperties(IReadOnlyDictionary<string, object?>? properties) {
        if (properties == null || properties.Count == 0) return default;

        var core = new Dictionary<(string Name, string Namespace), string>();
        Dictionary<(string Name, string Namespace), string>? application = null;
        foreach (var property in properties) {
            var element = ResolveStreamingProperty(property.Key);
            object? value = property.Value;
            if (element.Namespace == StreamingAppNamespace) {
                application ??= new() { [("Application", StreamingAppNamespace)] = "OfficeIMO.Excel" };
                application[element] = FormatStreamingProperty(value) ?? string.Empty;
            } else {
                string? text = value != null && (element.Namespace == StreamingTermsNamespace || element.Name == "lastPrinted")
                    ? FormatStreamingPropertyDate(value)
                    : FormatStreamingProperty(value);
                if (text == null) core.Remove(element);
                else core[element] = text;
            }
        }
        return (core.Count == 0 ? null : WriteStreamingPropertiesPart(core, "coreProperties", StreamingCoreNamespace),
            application == null ? null : WriteStreamingPropertiesPart(application, "Properties", StreamingAppNamespace));
    }

    private static string? FormatStreamingProperty(object? value) =>
        value == null ? null : OfficeRoundTripNumberText.FormatObject(value, CultureInfo.InvariantCulture);

    private static string FormatStreamingPropertyDate(object value) {
        DateTime date = value is DateTimeOffset offset ? offset.DateTime : System.Convert.ToDateTime(value, CultureInfo.InvariantCulture);
        return date.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture);
    }

    private static string WriteStreamingPropertiesPart(Dictionary<(string Name, string Namespace), string> properties,
        string root, string rootNamespace) {
        var buffer = new StringBuilder("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        using (XmlWriter writer = XmlWriter.Create(buffer, new XmlWriterSettings { OmitXmlDeclaration = true, NewLineChars = "\n" })) {
            writer.WriteStartElement(root, rootNamespace);
            if (rootNamespace == StreamingCoreNamespace) {
                writer.WriteAttributeString("xmlns", "dc", null, StreamingDcNamespace);
                writer.WriteAttributeString("xmlns", "dcterms", null, StreamingTermsNamespace);
                writer.WriteAttributeString("xmlns", "xsi", null, StreamingSchemaInstance);
            }
            foreach (var property in properties) {
                (string name, string ns) = property.Key;
                writer.WriteStartElement(name, ns);
                if (ns == StreamingTermsNamespace) writer.WriteAttributeString("xsi", "type", StreamingSchemaInstance, "dcterms:W3CDTF");
                writer.WriteString(property.Value);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }
        return buffer.ToString();
    }
}

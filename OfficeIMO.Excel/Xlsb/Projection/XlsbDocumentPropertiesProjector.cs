using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection;

/// <summary>Projects an XLSB package's core, extended and custom document properties into the editable workbook.</summary>
internal static class XlsbDocumentPropertiesProjector {
    private const string CoreNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private const string DublinCoreNamespace = "http://purl.org/dc/elements/1.1/";
    private const string DublinCoreTermsNamespace = "http://purl.org/dc/terms/";
    private const string ExtendedNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private const string CustomNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

    internal static void Apply(ExcelDocument document, XlsbWorkbook workbook) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(workbook);
        SpreadsheetDocument package = document._spreadSheetDocument;
        if (workbook.CorePropertiesXml != null) {
            if (TryLoadRoot(workbook.CorePropertiesXml, CoreNamespace, "coreProperties", out XElement? core)) {
                ApplyCore(package, core!);
            } else {
                Skip(workbook, "core");
            }
        }

        if (workbook.ExtendedPropertiesXml != null) {
            bool installed = TryLoadRoot(workbook.ExtendedPropertiesXml, ExtendedNamespace, "Properties", out _)
                && TryInstall(package, package.ExtendedFilePropertiesPart ?? package.AddExtendedFilePropertiesPart(),
                    workbook.ExtendedPropertiesXml, part => GC.KeepAlive(((ExtendedFilePropertiesPart)part).Properties));
            if (!installed) Skip(workbook, "extended");
        }

        if (workbook.CustomPropertiesXml != null) {
            bool installed = TryLoadRoot(workbook.CustomPropertiesXml, CustomNamespace, "Properties", out _)
                && TryInstall(package, package.CustomFilePropertiesPart ?? package.AddCustomFilePropertiesPart(),
                    workbook.CustomPropertiesXml, part => {
                        GC.KeepAlive(((CustomFilePropertiesPart)part).Properties);
                        document.LoadCustomDocumentProperties();
                    });
            if (!installed) {
                document.LoadCustomDocumentProperties();
                Skip(workbook, "custom");
            }
        }
    }

    // Installs the part and loads it through the typed API once, so a part that the SDK
    // cannot read is removed here instead of failing a later property read or save.
    private static bool TryInstall(SpreadsheetDocument package, OpenXmlPart part, byte[] bytes, Action<OpenXmlPart> verify) {
        try {
            using (var input = new MemoryStream(bytes, writable: false)) part.FeedData(input);
            verify(part);
            return true;
        } catch (Exception exception) when (exception is InvalidDataException or XmlException
            or OpenXmlPackageException or InvalidOperationException or FormatException) {
            package.DeletePart(part);
            return false;
        }
    }

    private static void Skip(XlsbWorkbook workbook, string kind) =>
        XlsbWorkbookReader.AddDocumentPropertiesSkipped(workbook, null,
            $"The {kind} document-property part is not well-formed property XML and was not imported.");

    private static void ApplyCore(SpreadsheetDocument package, XElement root) {
        var properties = package.PackageProperties;
        foreach (XElement element in root.Elements()) {
            string value = element.Value;
            switch (element.Name.NamespaceName) {
                case DublinCoreNamespace:
                    switch (element.Name.LocalName) {
                        case "title": properties.Title = value; break;
                        case "subject": properties.Subject = value; break;
                        case "creator": properties.Creator = value; break;
                        case "description": properties.Description = value; break;
                        case "identifier": properties.Identifier = value; break;
                        case "language": properties.Language = value; break;
                    }
                    break;
                case CoreNamespace:
                    switch (element.Name.LocalName) {
                        case "keywords": properties.Keywords = value; break;
                        case "lastModifiedBy": properties.LastModifiedBy = value; break;
                        case "revision": properties.Revision = value; break;
                        case "category": properties.Category = value; break;
                        case "version": properties.Version = value; break;
                        case "contentStatus": properties.ContentStatus = value; break;
                        case "lastPrinted": properties.LastPrinted = ParseDate(value) ?? properties.LastPrinted; break;
                    }
                    break;
                case DublinCoreTermsNamespace:
                    switch (element.Name.LocalName) {
                        case "created": properties.Created = ParseDate(value) ?? properties.Created; break;
                        case "modified": properties.Modified = ParseDate(value) ?? properties.Modified; break;
                    }
                    break;
            }
        }
    }

    private static DateTime? ParseDate(string value) {
        return DateTime.TryParse(value.Trim(), CultureInfo.InvariantCulture,
            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out DateTime parsed)
            ? DateTime.SpecifyKind(parsed, DateTimeKind.Utc)
            : null;
    }

    private static bool TryLoadRoot(byte[] bytes, string namespaceName, string localName, out XElement? root) {
        root = null;
        try {
            using var stream = new MemoryStream(bytes, writable: false);
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            });
            root = XDocument.Load(reader, LoadOptions.None).Root;
        } catch (XmlException) {
            // Malformed optional metadata stays out of the model and must not prevent loading.
            return false;
        }
        return root != null && root.Name == XName.Get(localName, namespaceName);
    }
}

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

    internal static void Apply(ExcelDocument document, XlsbWorkbook workbook) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(workbook);
        SpreadsheetDocument package = document._spreadSheetDocument;
        if (workbook.CorePropertiesXml != null && TryLoad(workbook.CorePropertiesXml, out XDocument? core)) {
            ApplyCore(package, core!.Root!);
        }

        if (workbook.ExtendedPropertiesXml != null && TryLoad(workbook.ExtendedPropertiesXml, out _)) {
            ExtendedFilePropertiesPart part = package.ExtendedFilePropertiesPart ?? package.AddExtendedFilePropertiesPart();
            using var input = new MemoryStream(workbook.ExtendedPropertiesXml, writable: false);
            part.FeedData(input);
        }

        if (workbook.CustomPropertiesXml != null && TryLoad(workbook.CustomPropertiesXml, out _)) {
            CustomFilePropertiesPart part = package.CustomFilePropertiesPart ?? package.AddCustomFilePropertiesPart();
            using var input = new MemoryStream(workbook.CustomPropertiesXml, writable: false);
            part.FeedData(input);
            document.LoadCustomDocumentProperties();
        }
    }

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

    private static bool TryLoad(byte[] bytes, out XDocument? document) {
        try {
            using var stream = new MemoryStream(bytes, writable: false);
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            });
            document = XDocument.Load(reader, LoadOptions.None);
            return document.Root != null;
        } catch (XmlException) {
            // Malformed optional metadata stays out of the model and must not prevent loading.
            document = null;
            return false;
        }
    }
}

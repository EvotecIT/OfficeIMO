using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;

namespace OfficeIMO.Excel {
    internal static class ExcelReadWorkbookParts {
        private const string OfficeRelationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix;
        private const string RichDataRelationship = "http://schemas.microsoft.com/office/2017/06/relationships/";

        internal const string WorksheetRelationship = OfficeRelationship + "worksheet";
        internal const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        internal const string SharedStringsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        internal const string StylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

        internal static readonly (string Relationship, string ContentType) SharedStrings = (OfficeRelationship + "sharedStrings", SharedStringsContentType);
        internal static readonly (string Relationship, string ContentType) Styles = (OfficeRelationship + "styles", StylesContentType);
        internal static readonly (string Relationship, string ContentType) Metadata = (OfficeRelationship + "sheetMetadata",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml");
        internal static readonly (string Relationship, string ContentType) RichValues = (RichDataRelationship + "rdRichValue",
            "application/vnd.ms-excel.rdrichvalue+xml");
        internal static readonly (string Relationship, string ContentType) RichStructures = (RichDataRelationship + "rdRichValueStructure",
            "application/vnd.ms-excel.rdrichvaluestructure+xml");

        internal static readonly HashSet<string> WorkbookContentTypes = new(StringComparer.OrdinalIgnoreCase) {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            "application/vnd.ms-excel.template.macroEnabled.main+xml",
            "application/vnd.ms-excel.addin.macroEnabled.main+xml"
        };

        internal static ExcelPartRef FindWorkbookPart(this ExcelPartRef workbook, (string Relationship, string ContentType) kind) {
            foreach ((_, ExcelPartRef part) in workbook.Parts) {
                if (part.RelationshipType is string type && OfficeXmlNamespaces.NormalizeRelationshipType(type) == kind.Relationship
                    && string.Equals(part.ContentType, kind.ContentType, StringComparison.OrdinalIgnoreCase)) return part;
            }
            return default;
        }

        internal static bool IsNonWorksheetRelationship(string? type) =>
            type == OfficeRelationship + "chartsheet" || type == OfficeRelationship + "dialogsheet"
            || type == OfficeRelationship + "macrosheet" || type == OfficeRelationship + "intlMacrosheet"
            || type == "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet"
            || type == "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";
    }
}

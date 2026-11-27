using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Workbook children in SpreadsheetML schema order, excluding extension content from lookups.</summary>
    internal static class XlsbWorkbookMarkup {
        private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
        private static readonly string[] ChildNames = {
            "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets",
            "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews",
            "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"
        };
        private static readonly OfficeXmlChildOrder ChildOrder = new(
            Enumerable.Repeat(Spreadsheet, ChildNames.Length).ToArray(), ChildNames);

        internal static OfficeXmlNode Root(ExcelDocument document) => ExcelPartRef.Workbook(document).Root;

        /// <summary>Finds a direct workbook child, without searching inside extensions or compatibility wrappers.</summary>
        internal static OfficeXmlNode Child(OfficeXmlNode workbook, string localName) => workbook.Child(Spreadsheet, localName);

        internal static OfficeXmlNode Add(OfficeXmlNode workbook, string localName) =>
            workbook.InsertOrdered(workbook.CreateElement(Spreadsheet, localName), ChildOrder);

        internal static string? UnsupportedAttribute(OfficeXmlNode element, string[] allowed) {
            foreach (OfficeXmlAttribute attribute in element.Attributes()) {
                if (attribute.NamespaceUri.Length != 0 || Array.IndexOf(allowed, attribute.LocalName) < 0) return attribute.LocalName;
            }
            return null;
        }

        // The typed SpreadsheetML values reject malformed attributes. A present but invalid value
        // must not become the default for an omitted attribute when native records are written.
        internal static bool? ReadBoolean(OfficeXmlNode element, string name) =>
            element.Attribute(name) is string value ? XmlConvert.ToBoolean(value) : null;

        internal static uint? ReadUInt32(OfficeXmlNode element, string name) =>
            element.Attribute(name) is string value ? XmlConvert.ToUInt32(value) : null;

        internal static int? ReadInt32(OfficeXmlNode element, string name) =>
            element.Attribute(name) is string value ? XmlConvert.ToInt32(value) : null;
    }
}

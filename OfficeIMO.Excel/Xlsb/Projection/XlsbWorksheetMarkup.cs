using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Handles SpreadsheetML worksheet children shared by projection and native save validation.</summary>
    internal static class XlsbWorksheetMarkup {
        private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;

        internal static OfficeXmlNode Root(ExcelSheet sheet) {
            OfficeXmlNode root = ExcelPartRef.Worksheet(sheet).Root;
            if (root.IsNone) throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            return root;
        }

        internal static OfficeXmlNode Child(OfficeXmlNode worksheet, string localName) =>
            worksheet.Child(Spreadsheet, localName);

        internal static OfficeXmlNode Create(OfficeXmlNode parent, string localName) =>
            parent.CreateElement(Spreadsheet, localName);

        internal static OfficeXmlNode Append(OfficeXmlNode parent, string localName) =>
            parent.Append(Create(parent, localName));

        /// <summary>Compares a projected child with the document's direct children of the same name.</summary>
        internal static bool MatchesChild(OfficeXmlNode worksheet, string localName, OfficeXmlNode expected) {
            var children = worksheet.Children(Spreadsheet, localName).GetEnumerator();
            if (!children.MoveNext()) return expected.IsNone;
            OfficeXmlNode actual = children.Current;
            return !expected.IsNone && !children.MoveNext() && XlsbProjectionXmlComparer.SameContent(actual, expected);
        }
    }
}

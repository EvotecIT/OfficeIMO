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
            return !expected.IsNone && !children.MoveNext() && SameContent(actual, expected);
        }

        // Keep traversal bounded by the supplied subtree and independent of the nesting depth.
        // Projection elements contain only element children or leaf text.
        private static bool SameContent(OfficeXmlNode actual, OfficeXmlNode expected) {
            // SDK text leaves can hide unknown child markup from the DOM traversal.
            if (actual.IsSdk && expected.IsSdk) return actual.OuterXml().AsSpan().SequenceEqual(expected.OuterXml());
            var actualNodes = actual;
            var expectedNodes = expected;
            while (actualNodes.Exists && expectedNodes.Exists) {
                if (!actualNodes.Is(expectedNodes.NamespaceUri, expectedNodes.LocalName)
                    || !SameAttributes(actualNodes.Attributes(), expectedNodes.Attributes())) return false;
                OfficeXmlNode actualChild = actualNodes.FirstChild;
                OfficeXmlNode expectedChild = expectedNodes.FirstChild;
                if (actualChild.Exists != expectedChild.Exists) return false;
                if (actualChild.Exists) {
                    actualNodes = actualChild;
                    expectedNodes = expectedChild;
                    continue;
                }
                if (!string.Equals(actualNodes.InnerText, expectedNodes.InnerText, StringComparison.Ordinal)) return false;

                while (true) {
                    if (actualNodes == actual || expectedNodes == expected) {
                        return actualNodes == actual && expectedNodes == expected
                            && string.Equals(actual.InnerText, expected.InnerText, StringComparison.Ordinal);
                    }
                    OfficeXmlNode actualNext = actualNodes.NextSibling;
                    OfficeXmlNode expectedNext = expectedNodes.NextSibling;
                    if (actualNext.Exists != expectedNext.Exists) return false;
                    if (actualNext.Exists) {
                        actualNodes = actualNext;
                        expectedNodes = expectedNext;
                        break;
                    }
                    actualNodes = actualNodes.Parent;
                    expectedNodes = expectedNodes.Parent;
                }
            }
            return false;
        }

        private static bool SameAttributes(IReadOnlyList<OfficeXmlAttribute> actual, IReadOnlyList<OfficeXmlAttribute> expected) {
            if (actual.Count != expected.Count) return false;
            foreach (OfficeXmlAttribute left in actual) {
                bool found = false;
                foreach (OfficeXmlAttribute right in expected) {
                    if (left.NamespaceUri == right.NamespaceUri && left.LocalName == right.LocalName && left.Value == right.Value) {
                        found = true;
                        break;
                    }
                }
                if (!found) return false;
            }
            return true;
        }
    }
}

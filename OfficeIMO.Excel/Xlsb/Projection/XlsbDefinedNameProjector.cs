using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects common XLSB defined names onto the workbook's SpreadsheetML names collection.</summary>
    internal static class XlsbDefinedNameProjector {
        internal const string CollectionName = "definedNames";
        internal const string ElementName = "definedName";
        private static readonly string[] NameAttributes = { "name", "localSheetId", "hidden", "comment" };

        internal static void Apply(OfficeXmlNode workbook, IReadOnlyList<XlsbDefinedName> sourceNames) {
            OfficeXmlNode names = default;
            foreach (XlsbDefinedName source in sourceNames) {
                if (string.IsNullOrWhiteSpace(source.FormulaText)) continue;
                if (names.IsNone) {
                    names = XlsbWorkbookMarkup.Child(workbook, CollectionName);
                    if (names.IsNone) names = XlsbWorkbookMarkup.Add(workbook, CollectionName);
                }
                OfficeXmlNode target = names.CreateElement(OfficeXmlNamespaces.SpreadsheetML, ElementName);
                target.SetAttribute("name", ToOpenXmlName(source));
                if (source.Hidden) target.SetBooleanAttribute("hidden", true);
                target.SetAttribute("comment", source.Comment ?? string.Empty);
                if (source.LocalSheetIndex != uint.MaxValue) target.SetUInt32Attribute("localSheetId", source.LocalSheetIndex);
                target.SetInnerText(source.FormulaText!);
                names.Append(target);
            }
        }

        internal static bool Matches(OfficeXmlNode actual, IReadOnlyList<XlsbDefinedName> expectedNames,
            IReadOnlyDictionary<string, string> reverseRenames) {
            int expectedIndex = 0;
            foreach (OfficeXmlNode item in Names(actual)) {
                while (expectedIndex < expectedNames.Count && string.IsNullOrWhiteSpace(expectedNames[expectedIndex].FormulaText)) {
                    expectedIndex++;
                }
                if (expectedIndex == expectedNames.Count) return false;
                XlsbDefinedName source = expectedNames[expectedIndex++];
                if (item.FirstChild.Exists
                    || UnsupportedAttribute(item) != null
                    || !string.Equals(item.Attribute("name"), ToOpenXmlName(source), StringComparison.Ordinal)
                    || !string.Equals(ExcelDocument.ReplaceSheetNameReferences(item.InnerText, reverseRenames), source.FormulaText, StringComparison.Ordinal)
                    || XlsbWorkbookMarkup.ReadUInt32(item, "localSheetId") != (source.LocalSheetIndex == uint.MaxValue ? (uint?)null : source.LocalSheetIndex)
                    || (XlsbWorkbookMarkup.ReadBoolean(item, "hidden") ?? false) != source.Hidden
                    || !string.Equals(item.Attribute("comment") ?? string.Empty, source.Comment ?? string.Empty, StringComparison.Ordinal)) {
                    return false;
                }
            }
            for (; expectedIndex < expectedNames.Count; expectedIndex++) {
                if (!string.IsNullOrWhiteSpace(expectedNames[expectedIndex].FormulaText)) return false;
            }
            return true;
        }

        internal static OfficeXmlNode.ChildList Names(OfficeXmlNode collection) =>
            collection.Children(OfficeXmlNamespaces.SpreadsheetML, ElementName);

        internal static string? UnsupportedAttribute(OfficeXmlNode name) =>
            XlsbWorkbookMarkup.UnsupportedAttribute(name, NameAttributes);

        /// <summary>Reflects native reference normalization after a successful write, including references to deleted sheets.</summary>
        internal static void SynchronizeExisting(OfficeXmlNode workbook, IReadOnlyList<XlsbDefinedName> names) {
            foreach (OfficeXmlNode current in Names(XlsbWorkbookMarkup.Child(workbook, CollectionName))) {
                string? name = current.Attribute("name");
                uint scope = XlsbWorkbookMarkup.ReadUInt32(current, "localSheetId") ?? uint.MaxValue;
                XlsbDefinedName? source = names.FirstOrDefault(item => ToOpenXmlName(item) == name && item.LocalSheetIndex == scope);
                if (source?.FormulaText != null) current.SetInnerText(source.FormulaText);
            }
        }

        private static string ToOpenXmlName(XlsbDefinedName source) {
            if (!source.BuiltIn) return source.Name;
            if (string.Equals(source.Name, "Print_Area", StringComparison.OrdinalIgnoreCase)) return "_xlnm.Print_Area";
            if (string.Equals(source.Name, "Print_Titles", StringComparison.OrdinalIgnoreCase)) return "_xlnm.Print_Titles";
            return source.Name;
        }
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects common XLSB defined names onto the normal workbook model.</summary>
    internal static class XlsbDefinedNameProjector {
        internal static void Apply(ExcelDocument document, IReadOnlyList<XlsbDefinedName> sourceNames) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (sourceNames == null) throw new ArgumentNullException(nameof(sourceNames));

            DefinedNames? targetNames = null;
            foreach (XlsbDefinedName source in sourceNames) {
                if (string.IsNullOrWhiteSpace(source.FormulaText)) continue;
                targetNames ??= GetOrCreateDefinedNames(document.WorkbookRoot);
                var target = new DefinedName {
                    Name = ToOpenXmlName(source),
                    Text = source.FormulaText!,
                    Hidden = source.Hidden ? true : (bool?)null,
                    Comment = source.Comment
                };
                if (source.LocalSheetIndex != uint.MaxValue) {
                    target.LocalSheetId = source.LocalSheetIndex;
                }
                targetNames.Append(target);
            }
        }

        internal static bool Matches(DefinedNames? actual, IReadOnlyList<XlsbDefinedName> expectedNames) {
            if (expectedNames == null) throw new ArgumentNullException(nameof(expectedNames));
            XlsbDefinedName[] expected = expectedNames
                .Where(name => !string.IsNullOrWhiteSpace(name.FormulaText))
                .ToArray();
            DefinedName[] current = actual?.Elements<DefinedName>().ToArray() ?? Array.Empty<DefinedName>();
            if (current.Length != expected.Length) return false;

            for (int index = 0; index < current.Length; index++) {
                DefinedName item = current[index];
                XlsbDefinedName source = expected[index];
                if (item.HasChildren
                    || HasUnsupportedAttributes(item)
                    || !string.Equals(item.Name?.Value, ToOpenXmlName(source), StringComparison.Ordinal)
                    || !string.Equals(item.Text ?? string.Empty, source.FormulaText ?? string.Empty, StringComparison.Ordinal)
                    || item.LocalSheetId?.Value != (source.LocalSheetIndex == uint.MaxValue ? (uint?)null : source.LocalSheetIndex)
                    || (item.Hidden?.Value ?? false) != source.Hidden
                    || !string.Equals(item.Comment?.Value, source.Comment, StringComparison.Ordinal)) {
                    return false;
                }
            }
            return true;
        }

        private static DefinedNames GetOrCreateDefinedNames(Workbook workbook) {
            DefinedNames? existing = workbook.GetFirstChild<DefinedNames>();
            if (existing != null) return existing;

            var definedNames = new DefinedNames();
            OpenXmlElement? before = workbook.GetFirstChild<CalculationProperties>();
            if (before != null) {
                workbook.InsertBefore(definedNames, before);
            } else {
                workbook.Append(definedNames);
            }
            return definedNames;
        }

        private static string ToOpenXmlName(XlsbDefinedName source) {
            if (!source.BuiltIn) return source.Name;
            if (string.Equals(source.Name, "Print_Area", StringComparison.OrdinalIgnoreCase)) return "_xlnm.Print_Area";
            if (string.Equals(source.Name, "Print_Titles", StringComparison.OrdinalIgnoreCase)) return "_xlnm.Print_Titles";
            return source.Name;
        }

        private static bool HasUnsupportedAttributes(DefinedName name) {
            foreach (OpenXmlAttribute attribute in name.GetAttributes()) {
                if (string.Equals(attribute.NamespaceUri, "http://www.w3.org/2000/xmlns/", StringComparison.Ordinal)) continue;
                if (attribute.LocalName != "name"
                    && attribute.LocalName != "localSheetId"
                    && attribute.LocalName != "hidden"
                    && attribute.LocalName != "comment") {
                    return true;
                }
            }
            return false;
        }
    }
}

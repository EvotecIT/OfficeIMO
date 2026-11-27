using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Builds the editable projection of a BIFF12 workbook directly as an OfficeIMO Open XML package.</summary>
    internal static class XlsbOwnedWorkbookProjector {
        private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
        private const int Relationships = OfficeXmlNamespaces.Relationships;
        private const string RelationshipPrefix = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix;
        private const string WorkbookPart = "/xl/workbook.xml";
        private const string StylesPart = "/xl/styles.xml";
        private const string ThemePart = "/xl/theme/theme1.xml";
        private const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        private const string ChartsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

        // The legacy table and comment projectors and the initial custom-property load still use their own model paths.
        // Their native contents must be preserved until those paths can run against the owned package.
        internal static bool CanWrite(XlsbWorkbook workbook) =>
            workbook.Sheets.Count != 0 && workbook.CustomPropertiesXml == null
            && workbook.Sheets.All(sheet => sheet.Tables.Count == 0 && sheet.Comments == null
                && XlsbWorksheetGridProjector.CanStream(sheet));

        internal static byte[] Write(OfficeOpcEditablePackage package, XlsbWorkbook native,
            Func<string, ISet<string>, string> worksheetName, Action<string, string?, bool> recordMutation,
            CancellationToken cancellationToken) {
            OfficeXmlTree workbookTree = package.GetXml(WorkbookPart);
            var workbook = new OfficeXmlNode(workbookTree, workbookTree.Root);
            var related = new XlsbOwnedRelatedParts(package, native, cancellationToken);
            if (native.Stylesheet != null) {
                OfficeXmlTree styles = package.GetXml(StylesPart);
                var root = new OfficeXmlNode(styles, styles.Root);
                for (OfficeXmlNode child = root.FirstChild; child.Exists; child = root.FirstChild) child.Remove();
                XlsbStylesheetMarkup.Write(root, native.Stylesheet);
            }

            if (native.Uses1904DateSystem || !string.IsNullOrEmpty(native.CodeName)) {
                OfficeXmlNode properties = XlsbWorkbookMarkup.Add(workbook, "workbookPr");
                if (native.Uses1904DateSystem) properties.SetBooleanAttribute("date1904", true);
                if (!string.IsNullOrEmpty(native.CodeName)) properties.SetAttribute("codeName", native.CodeName);
            }
            if (native.WorkbookProtection != null) XlsbWorkbookProtectionProjector.Apply(workbook, native.WorkbookProtection);

            int originalActive = (int)Math.Min(native.ActiveTab ?? 0U, (uint)(native.Sheets.Count - 1));
            int firstSheet = (int)Math.Min(native.FirstSheet ?? 0U, (uint)(native.Sheets.Count - 1));
            int active = ChooseVisibleTab(native.Sheets, originalActive);
            if (active != originalActive && firstSheet > active) firstSheet = active;
            if (active != 0 || originalActive != 0 || firstSheet != 0) {
                OfficeXmlNode view = XlsbWorksheetMarkup.Append(XlsbWorkbookMarkup.Add(workbook, "bookViews"), "workbookView");
                view.SetUInt32Attribute("activeTab", (uint)active);
                if (firstSheet != 0) view.SetUInt32Attribute("firstSheet", (uint)firstSheet);
            }
            OfficeXmlNode sheets = XlsbWorkbookMarkup.Add(workbook, "sheets");

            ExcelDeferredSharedStrings strings = ExcelDeferredSharedStrings.Load(package, partName: null)!;
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int worksheetIndex = 0, chartIndex = 0, tabIndex = 0;
            foreach (XlsbSheet source in native.Sheets) {
                cancellationToken.ThrowIfCancellationRequested();
                string name = source.IsChartSheet ? source.Name : worksheetName(source.Name, names);
                names.Add(name);
                bool selected = active != originalActive && tabIndex == active
                    || (active == originalActive || tabIndex != originalActive) && source.TabSelected == true;
                string relativeName;
                string relationshipType;
                if (source.IsChartSheet) {
                    relativeName = "chartsheets/sheet" + (++chartIndex).ToString(CultureInfo.InvariantCulture) + ".xml";
                    relationshipType = RelationshipPrefix + "chartsheet";
                    string nameInPackage = "/xl/" + relativeName;
                    OfficeXmlTree tree = package.CreateXmlPart(nameInPackage, ChartsheetContentType, Spreadsheet, "chartsheet");
                    XlsbOwnedWorksheetMarkup.WriteChartSheet(new OfficeXmlNode(tree, tree.Root), source, selected, nameInPackage, related);
                } else {
                    relativeName = "worksheets/sheet" + (++worksheetIndex).ToString(CultureInfo.InvariantCulture) + ".xml";
                    relationshipType = RelationshipPrefix + "worksheet";
                    string nameInPackage = "/xl/" + relativeName;
                    OfficeXmlNode layout = XlsbOwnedWorksheetMarkup.Create(package, nameInPackage, source, selected, related);
                    foreach (XlsbCell cell in source.Cells) {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (cell.Kind == XlsbCellValueKind.Text && cell.Value is string { Length: > 0 } text
                            && strings.TryGetOrAdd(text, ExcelSheet.CellValuePlainStringPromotionSharedStringCount, out _)) {
                            strings.AddedReferences++;
                        }
                    }
                    package.CreateStreamedPart(nameInPackage, WorksheetContentType, stream =>
                        XlsbWorksheetGridProjector.WriteWorksheet(stream, source, LookupSharedString,
                            (reference, retainsCache) => recordMutation(nameInPackage, reference, retainsCache),
                            cancellationToken, layout, native.Stylesheet));
                }
                OfficeXmlNode sheet = XlsbWorksheetMarkup.Append(sheets, "sheet");
                sheet.SetAttribute("name", name);
                sheet.SetUInt32Attribute("sheetId", checked((uint)++tabIndex));
                sheet.SetAttribute(Relationships, "id", package.AddRelationship(WorkbookPart, relationshipType, relativeName));
                if (source.State == 1) sheet.SetAttribute("state", "hidden");
                else if (source.State == 2) sheet.SetAttribute("state", "veryHidden");
            }

            XlsbDefinedNameProjector.Apply(workbook, native.DefinedNames);
            if (native.CalculationProperties != null) XlsbCalculationPropertiesProjector.Apply(workbook, native.CalculationProperties);
            if (native.ThemePartName != null) related.ReplaceTheme(WorkbookPart, ThemePart, native.ThemePartName);
            if (native.VbaProjectPartName != null)
                related.Link(WorkbookPart, native.VbaProjectPartName, "http://schemas.microsoft.com/office/2006/relationships/vbaProject");
            if (!strings.TryWrite(package, WorkbookPart, declareNewCounts: false)) throw new InvalidOperationException("The projected workbook cannot add its shared strings.");

            XlsbDocumentPropertiesProjector.Apply(package, native);
            using var output = new MemoryStream(16 * 1024);
            package.Save(output);
            return output.ToArray();

            int? LookupSharedString(string text) {
                int index = strings.IndexOf(text);
                return index < 0 ? null : index;
            }
        }

        private static int ChooseVisibleTab(IReadOnlyList<XlsbSheet> sheets, int active) {
            if (sheets[active].State is not (1 or 2)) return active;
            for (int index = active + 1; index < sheets.Count; index++) {
                if (sheets[index].State is not (1 or 2)) return index;
            }
            for (int index = active - 1; index >= 0; index--) {
                if (sheets[index].State is not (1 or 2)) return index;
            }
            return active;
        }
    }
}

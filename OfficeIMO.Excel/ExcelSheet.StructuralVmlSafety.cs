using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private void ValidateStructuralVmlControlSafety(MutationPlanScanBudget? budget = null) {
            IEnumerable<VmlDrawingPart> workbookVmlParts =
                WorkbookPartRoot.WorksheetParts.SelectMany(part => part.VmlDrawingParts)
                    .Concat(WorkbookPartRoot.DialogsheetParts.SelectMany(part => part.VmlDrawingParts))
                    .Concat(WorkbookPartRoot.ChartsheetParts.SelectMany(part => part.VmlDrawingParts))
                    .Distinct();
            bool hasUnsupportedControl = false;
            foreach (WorksheetPart worksheetPart in InspectMutationPlanElements(WorkbookPartRoot.WorksheetParts, budget)) {
                if (InspectMutationPlanElements(
                        worksheetPart.Worksheet?.Descendants<Controls>() ?? Enumerable.Empty<Controls>(),
                        budget).Any()
                    || InspectMutationPlanElements(worksheetPart.ControlPropertiesParts, budget).Any()) {
                    hasUnsupportedControl = true;
                    break;
                }
            }
            if (hasUnsupportedControl || ContainsUnsupportedVmlFormControl(workbookVmlParts, budget)) {
                throw new InvalidOperationException(
                    "Cannot structurally edit a workbook containing form controls because their anchors and cross-sheet links cannot yet be remapped safely.");
            }
            if (InspectMutationPlanElements(WorksheetRoot.Descendants<OleObjects>(), budget).Any()
                || InspectMutationPlanElements(_worksheetPart.EmbeddedObjectParts, budget).Any()) {
                throw new InvalidOperationException(
                    "Cannot structurally edit a worksheet containing embedded OLE objects because their VML anchors cannot yet be remapped safely.");
            }
            if (_worksheetPart.SingleCellTablePart != null) {
                budget?.Consume();
                throw new InvalidOperationException(
                    "Cannot edit rows on a worksheet containing single-cell XML mappings because their mapped references cannot yet be remapped safely.");
            }
            if (InspectMutationPlanElements(WorkbookPartRoot.MacroSheetParts, budget).Any()
                || InspectMutationPlanElements(WorkbookPartRoot.InternationalMacroSheetParts, budget).Any()) {
                throw new InvalidOperationException(
                    "Cannot edit rows in a workbook containing Excel 4.0 macro sheets because their formulas cannot yet be remapped safely.");
            }
            if (WorkbookPartRoot.WorkbookRevisionHeaderPart != null) {
                budget?.Consume();
                throw new InvalidOperationException(
                    "Cannot edit rows while legacy workbook revision tracking is present because revision-log references cannot yet be remapped safely.");
            }
        }

        private bool ContainsUnsupportedVmlFormControl(
            IEnumerable<VmlDrawingPart> vmlParts,
            MutationPlanScanBudget? budget = null) {
            XNamespace excelNamespace = "urn:schemas-microsoft-com:office:excel";
            foreach (VmlDrawingPart vmlPart in InspectMutationPlanElements(vmlParts, budget)) {
                XDocument document = LoadOrCreateVmlDocument(vmlPart);
                foreach (XElement clientData in InspectMutationPlanElements(document.Descendants(), budget)
                    .Where(element => element.Name == excelNamespace + "ClientData")) {
                    string? objectType = clientData.Attribute("ObjectType")?.Value;
                    if (!string.Equals(objectType, "Note", StringComparison.OrdinalIgnoreCase)) return true;
                }
            }
            return false;
        }
    }
}

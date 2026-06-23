using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal void CleanupCalculationArtifacts(bool save = true, ExcelCalculationCleanupPolicy policy = ExcelCalculationCleanupPolicy.PreserveExistingCalculationProperties) {
            var workbookPart = WorkbookPartRoot;

            foreach (var calculationChainPart in workbookPart.GetPartsOfType<CalculationChainPart>().ToList()) {
                workbookPart.DeletePart(calculationChainPart);
            }

            bool hasCellFormulas = workbookPart.WorksheetParts
                .Any(worksheetPart => (worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is null."))
                    .Descendants<CellFormula>()
                    .Any(formula => !string.IsNullOrWhiteSpace(formula.Text)));

            var workbook = WorkbookRoot;
            var calculationProperties = workbook.Elements<CalculationProperties>().FirstOrDefault();
            if (!hasCellFormulas) {
                if (save) {
                    workbook.Save();
                }
                return;
            }

            if (policy == ExcelCalculationCleanupPolicy.RequestFullCalculationOnOpen) {
                if (calculationProperties == null) {
                    calculationProperties = new CalculationProperties();
                    InsertCalculationPropertiesInSchemaOrder(workbook, calculationProperties);
                }

                calculationProperties.SetAttribute(new OpenXmlAttribute("calcId", string.Empty, "191029"));
                calculationProperties.SetAttribute(new OpenXmlAttribute("fullCalcOnLoad", string.Empty, "1"));
                calculationProperties.SetAttribute(new OpenXmlAttribute("forceFullCalc", string.Empty, "1"));
            } else if (policy == ExcelCalculationCleanupPolicy.ClearAutomaticFullCalculationOnOpen && calculationProperties != null) {
                calculationProperties.RemoveAttribute("fullCalcOnLoad", string.Empty);
                calculationProperties.RemoveAttribute("forceFullCalc", string.Empty);
            }

            if (save) {
                workbook.Save();
            }
        }
    }
}

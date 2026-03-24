using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        internal void CleanupCalculationArtifacts(bool save = true) {
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

            if (calculationProperties == null) {
                calculationProperties = new CalculationProperties();
                var extensionList = workbook.Elements<ExtensionList>().FirstOrDefault();
                if (extensionList != null) {
                    workbook.InsertBefore(calculationProperties, extensionList);
                } else {
                    workbook.Append(calculationProperties);
                }
            }

            calculationProperties.SetAttribute(new OpenXmlAttribute("calcId", string.Empty, "191029"));
            calculationProperties.SetAttribute(new OpenXmlAttribute("fullCalcOnLoad", string.Empty, "1"));
            calculationProperties.SetAttribute(new OpenXmlAttribute("forceFullCalc", string.Empty, "1"));

            if (save) {
                workbook.Save();
            }
        }
    }
}

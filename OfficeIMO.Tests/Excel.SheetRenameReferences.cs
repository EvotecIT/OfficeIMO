using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using TableColumn = DocumentFormat.OpenXml.Spreadsheet.TableColumn;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_RenameWorkSheet_UpdatesFormulasDefinedNamesAndInternalHyperlinks() {
            string filePath = Path.Combine(_directoryWithFiles, "RenameWorksheet.References.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data", SheetNameValidationMode.Strict);
                    var summary = document.AddWorkSheet("Summary", SheetNameValidationMode.Strict);

                    data.CellValue(1, 1, 42);
                    summary.CellFormula(1, 1, "SUM(Data!A1)");
                    summary.SetInternalLink(2, 1, data, "A1", display: "Go");

                    document.SetNamedRange("GlobalData", "'Data'!A1:B2", save: false);
                    document.SetPrintArea(data, "A1:B2", save: false);
                    document.SetPrintTitles(data, firstRow: 1, lastRow: 1, firstCol: 1, lastCol: 1, save: false);

                    data.Name = "Renamed";
                    document.Save(filePath, openExcel: false);
                }

                using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
                var workbookPart = spreadsheet.WorkbookPart!;

                var definedNames = workbookPart.Workbook.DefinedNames!.Elements<DefinedName>().ToList();
                Assert.Contains(definedNames, dn => dn.Name == "GlobalData" && dn.Text == "'Renamed'!$A$1:$B$2");
                Assert.Contains(definedNames, dn => dn.Name == "_xlnm.Print_Area" && dn.Text == "'Renamed'!$A$1:$B$2");
                Assert.Contains(definedNames, dn => dn.Name == "_xlnm.Print_Titles" && dn.Text == "'Renamed'!$1:$1,'Renamed'!$A:$A");

                var summarySheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name == "Summary");
                var summaryPart = (WorksheetPart)workbookPart.GetPartById(summarySheet.Id!);

                var formulaCell = summaryPart.Worksheet.Descendants<Cell>()
                    .First(c => string.Equals(c.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase));
                Assert.Equal("SUM('Renamed'!A1)", formulaCell.CellFormula!.Text);

                var hyperlink = summaryPart.Worksheet.Descendants<Hyperlink>()
                    .First(h => string.Equals(h.Reference?.Value, "A2", StringComparison.OrdinalIgnoreCase));
                Assert.Equal("'Renamed'!A1", hyperlink.Location?.Value);
            }
            finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_RenameWorkSheet_UpdatesChartsPivotsAndSparklines() {
            string filePath = Path.Combine(_directoryWithFiles, "RenameWorksheet.ChartPivotSparkline.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data", SheetNameValidationMode.Strict);

                    data.CellValue(1, 1, "Month");
                    data.CellValue(1, 2, "Sales");
                    data.CellValue(2, 1, "Jan");
                    data.CellValue(2, 2, 10);
                    data.CellValue(3, 1, "Feb");
                    data.CellValue(3, 2, 20);

                    data.AddChartFromRange("A1:B3", row: 1, column: 4, widthPixels: 320, heightPixels: 220);
                    data.AddPivotTable("A1:B3", "F1");
                    data.AddSparklines("'Data'!B2:B3", "C2:C3");

                    data.Name = "Renamed Data";
                    document.Save(filePath, openExcel: false);
                }

                using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
                var workbookPart = spreadsheet.WorkbookPart!;
                var renamedSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name == "Renamed Data");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(renamedSheet.Id!);

                var sparklineFormulas = worksheetPart.Worksheet.Descendants<OfficeFormula>()
                    .Select(f => f.Text)
                    .Where(f => !string.IsNullOrWhiteSpace(f))
                    .ToList();
                Assert.Contains("'Renamed Data'!B2:B3", sparklineFormulas);

                var chartPart = worksheetPart.DrawingsPart!.ChartParts.First();
                var chartFormulas = chartPart.ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>()
                    .Select(f => f.Text)
                    .Where(f => !string.IsNullOrWhiteSpace(f))
                    .ToList();
                Assert.NotEmpty(chartFormulas);
                Assert.Contains(chartFormulas, f => f!.Contains("'Renamed Data'!", StringComparison.Ordinal));
                Assert.DoesNotContain(chartFormulas, f => f!.Contains("'Data'!", StringComparison.Ordinal));

                var pivotCache = workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>().Single();
                Assert.Equal("Renamed Data", pivotCache.PivotCacheDefinition!.CacheSource!.WorksheetSource!.Sheet!.Value);
            }
            finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_RenameWorkSheet_UpdatesDataValidationAndConditionalFormattingFormulas() {
            string filePath = Path.Combine(_directoryWithFiles, "RenameWorksheet.ValidationAndCf.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data", SheetNameValidationMode.Strict);
                    var summary = document.AddWorkSheet("Summary", SheetNameValidationMode.Strict);

                    data.CellValue(1, 1, 5);
                    data.CellValue(2, 1, 15);

                    summary.ValidationCustomFormula("A1:A2", "COUNTIF(Data!$A$1:$A$2,\">0\")>0");
                    summary.AddConditionalRule("B1:B2", ConditionalFormattingOperatorValues.GreaterThan, "Data!$A$1");

                    data.Name = "Renamed Data";
                    document.Save(filePath, openExcel: false);
                }

                using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
                var workbookPart = spreadsheet.WorkbookPart!;
                var summarySheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name == "Summary");
                var summaryPart = (WorksheetPart)workbookPart.GetPartById(summarySheet.Id!);

                var validation = summaryPart.Worksheet.Descendants<DataValidation>().Single();
                Assert.Equal("COUNTIF('Renamed Data'!$A$1:$A$2,\">0\")>0", validation.GetFirstChild<Formula1>()!.Text);

                var rule = summaryPart.Worksheet.Descendants<ConditionalFormattingRule>().Single();
                Assert.Equal("'Renamed Data'!$A$1", rule.Elements<Formula>().Single().Text);
            }
            finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_RenameWorkSheet_RefreshesTocAndBacklinkDisplayText() {
            string filePath = Path.Combine(_directoryWithFiles, "RenameWorksheet.TocAndBacklinks.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data", SheetNameValidationMode.Strict);
                    var summary = document.AddWorkSheet("Summary", SheetNameValidationMode.Strict);

                    data.CellValue(1, 1, "Value");
                    summary.CellValue(1, 1, "Other");

                    document.AddTableOfContents();
                    document.AddBackLinksToToc();

                    data.Name = "Renamed Data";
                    var toc = document.Sheets.First(s => s.Name == "TOC");
                    toc.Name = "Index";

                    document.Save(filePath, openExcel: false);
                }

                using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
                var workbookPart = spreadsheet.WorkbookPart!;

                var indexSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name == "Index");
                var indexPart = (WorksheetPart)workbookPart.GetPartById(indexSheet.Id!);
                var tocLink = indexPart.Worksheet.Descendants<Hyperlink>()
                    .First(h => h.Location?.Value == "'Renamed Data'!A1");
                Assert.Equal("Renamed Data", GetCellText(workbookPart, indexPart, tocLink.Reference!.Value!));

                var renamedSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name == "Renamed Data");
                var renamedPart = (WorksheetPart)workbookPart.GetPartById(renamedSheet.Id!);
                var backlink = renamedPart.Worksheet.Descendants<Hyperlink>()
                    .First(h => string.Equals(h.Reference?.Value, "A2", StringComparison.OrdinalIgnoreCase));
                Assert.Equal("'Index'!A1", backlink.Location?.Value);
                Assert.Equal("← Index", GetCellText(workbookPart, renamedPart, "A2"));
            }
            finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_RenameWorkSheet_PreservesExternalWorkbookReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "RenameWorksheet.ExternalReferences.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data", SheetNameValidationMode.Strict);
                    var summary = document.AddWorkSheet("Summary", SheetNameValidationMode.Strict);

                    data.CellValue(1, 1, 10);
                    summary.CellFormula(1, 1, "SUM(Data!A1,'[Other.xlsx]Data'!A1,[Other.xlsx]Data!A1)");
                    summary.ValidationCustomFormula("A2:A3", "COUNTIF(Data!$A$1:$A$3,\">0\")+COUNTIF('[Other.xlsx]Data'!$A$1:$A$3,\">0\")>0");

                    data.Name = "Renamed Data";
                    document.Save(filePath, openExcel: false);
                }

                using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
                var workbookPart = spreadsheet.WorkbookPart!;
                var summarySheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name == "Summary");
                var summaryPart = (WorksheetPart)workbookPart.GetPartById(summarySheet.Id!);

                var formulaCell = summaryPart.Worksheet.Descendants<Cell>()
                    .First(c => string.Equals(c.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase));
                Assert.Equal("SUM('Renamed Data'!A1,'[Other.xlsx]Data'!A1,[Other.xlsx]Data!A1)", formulaCell.CellFormula!.Text);

                var validation = summaryPart.Worksheet.Descendants<DataValidation>().Single();
                Assert.Equal("COUNTIF('Renamed Data'!$A$1:$A$3,\">0\")+COUNTIF('[Other.xlsx]Data'!$A$1:$A$3,\">0\")>0", validation.GetFirstChild<Formula1>()!.Text);
            }
            finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_RenameWorkSheet_UpdatesTableDefinitionFormulas() {
            string filePath = Path.Combine(_directoryWithFiles, "RenameWorksheet.TableFormulas.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data", SheetNameValidationMode.Strict);
                    var summary = document.AddWorkSheet("Summary", SheetNameValidationMode.Strict);

                    data.CellValue(1, 1, 10);
                    data.CellValue(2, 1, 20);

                    summary.CellValue(1, 1, "Label");
                    summary.CellValue(1, 2, "Value");
                    summary.CellValue(2, 1, "A");
                    summary.CellValue(2, 2, 1);
                    summary.CellValue(3, 1, "B");
                    summary.CellValue(3, 2, 2);
                    summary.AddTable("A1:B3", true, "SummaryTable", OfficeIMO.Excel.TableStyle.TableStyleMedium9);

                    var workbookPart = document._spreadSheetDocument.WorkbookPart!;
                    var summarySheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name == "Summary");
                    var summaryPart = (WorksheetPart)workbookPart.GetPartById(summarySheet.Id!);
                    var tablePart = summaryPart.TableDefinitionParts.Single();
                    var valueColumn = tablePart.Table.TableColumns!.Elements<TableColumn>().Last();
                    valueColumn.CalculatedColumnFormula = new CalculatedColumnFormula { Text = "SUM(Data!$A$2,1)" };
                    valueColumn.TotalsRowFormula = new TotalsRowFormula { Text = "SUM(Data!$A$2:$A$3)" };
                    tablePart.Table.TotalsRowShown = true;
                    tablePart.Table.Save();

                    data.Name = "Renamed Data";
                    document.Save(filePath, openExcel: false);
                }

                using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
                var workbookPartAfter = spreadsheet.WorkbookPart!;
                var summarySheetAfter = workbookPartAfter.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name == "Summary");
                var summaryPartAfter = (WorksheetPart)workbookPartAfter.GetPartById(summarySheetAfter.Id!);
                var tablePartAfter = summaryPartAfter.TableDefinitionParts.Single();
                var valueColumnAfter = tablePartAfter.Table.TableColumns!.Elements<TableColumn>().Last();

                Assert.Equal("SUM('Renamed Data'!$A$2,1)", valueColumnAfter.CalculatedColumnFormula!.Text);
                Assert.Equal("SUM('Renamed Data'!$A$2:$A$3)", valueColumnAfter.TotalsRowFormula!.Text);
            }
            finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
        private static string GetCellText(WorkbookPart workbookPart, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<Cell>()
                .First(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));

            if (cell.DataType?.Value == CellValues.SharedString) {
                int id = int.Parse(cell.CellValue!.InnerText, System.Globalization.CultureInfo.InvariantCulture);
                var item = workbookPart.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>().ElementAt(id);
                if (item.Text != null) {
                    return item.Text.Text ?? string.Empty;
                }

                return string.Concat(item.Descendants<Text>().Select(t => t.Text));
            }

            if (cell.DataType?.Value == CellValues.InlineString) {
                return cell.InlineString?.Text?.Text ?? string.Concat(cell.InlineString?.Descendants<Text>().Select(t => t.Text) ?? Enumerable.Empty<string>());
            }

            return cell.CellValue?.InnerText ?? string.Empty;
        }
    }
}

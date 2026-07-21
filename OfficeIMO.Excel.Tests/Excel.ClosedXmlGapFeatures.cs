using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using ExcelTableStyle = OfficeIMO.Excel.TableStyle;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ClosedXmlGap_ObjectModel_RichText_Sort_Table_And_Clear() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ObjectModel.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");

                sheet.CellAt(1, 1).SetValue("Name").SetBold().SetFillColor("FFF2CC");
                sheet.CellAt(1, 2).SetValue("Score").SetBold().SetFillColor("FFF2CC");
                sheet.CellAt(2, 1).SetValue("Charlie");
                sheet.CellAt(2, 2).SetValue(30d);
                sheet.CellAt(3, 1).SetValue("Alice");
                sheet.CellAt(3, 2).SetValue(10d);
                sheet.CellAt(4, 1).SetValue("Bob");
                sheet.CellAt(4, 2).SetValue(20d);
                sheet.SetComment(3, 1, "temporary comment", author: "Tester", initials: "TT");

                sheet.Range("A1:B4").SortByColumn(2, ascending: true, hasHeader: true);
                Assert.Equal("Alice", sheet.CellAt(2, 1).GetValue<string>());
                Assert.Equal(10d, sheet.CellAt(2, 2).GetValue<double>());
                Assert.True(sheet.HasComment(2, 1));
                Assert.False(sheet.HasComment(3, 1));

                sheet.Range("A1:B4").ApplyAutoFilter();
                ExcelTable table = sheet.Range("A1:B4").CreateTable("Scores");
                Assert.Equal("A1:B4", table.Range);
                table.SetStyle(ExcelTableStyle.TableStyleMedium4);

                sheet.Range("D1:E1").Merge().Unmerge();
                sheet.CellAt(6, 1).SetRichText(
                    new ExcelRichTextRun("Strong") { Bold = true, FontColor = "FF0000" },
                    ExcelRichTextRun.Plain(" text"));

                IReadOnlyList<ExcelRichTextRun> runs = sheet.CellAt(6, 1).GetRichText();
                Assert.Equal(2, runs.Count);
                Assert.True(runs[0].Bold);
                Assert.Equal("Strong", runs[0].Text);

                sheet.ClearRange("A2:A2", ExcelClearOptions.Values | ExcelClearOptions.Comments);
                Assert.True(sheet.CellAt(2, 1).GetValue().IsBlank);
                Assert.False(sheet.HasComment(2, 1));

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_MergeRange_DuplicateCallKeepsSingleMerge() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.MergeDuplicate.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Merge");
                sheet.MergeRange("A1:B1");
                sheet.MergeRange("A1:B1");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                MergeCell merge = Assert.Single(worksheetPart.Worksheet.Descendants<MergeCell>());
                Assert.Equal("A1:B1", merge.Reference?.Value);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_RangeFormatting_ReusesStylesAndValidates() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.RangeFormatting.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, 10d);
                sheet.CellValue(1, 2, 20d);
                sheet.CellValue(2, 1, 30d);
                sheet.CellValue(2, 2, 40d);

                sheet.Range("A1:B2").SetNumberFormat("0.00");
                sheet.Range("C1:D2").SetFillColor("#FFF2CC");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorkbookStylesPart stylesPart = spreadsheet.WorkbookPart!.WorkbookStylesPart!;
                Stylesheet stylesheet = stylesPart.Stylesheet!;
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value != null)
                    .ToDictionary(cell => cell.CellReference!.Value!);

                uint numberStyleIndex = cells["A1"].StyleIndex!.Value;
                Assert.NotEqual(0U, numberStyleIndex);
                Assert.Equal(numberStyleIndex, cells["A2"].StyleIndex!.Value);
                Assert.Equal(numberStyleIndex, cells["B1"].StyleIndex!.Value);
                Assert.Equal(numberStyleIndex, cells["B2"].StyleIndex!.Value);

                uint fillStyleIndex = cells["C1"].StyleIndex!.Value;
                Assert.NotEqual(0U, fillStyleIndex);
                Assert.Equal(fillStyleIndex, cells["C2"].StyleIndex!.Value);
                Assert.Equal(fillStyleIndex, cells["D1"].StyleIndex!.Value);
                Assert.Equal(fillStyleIndex, cells["D2"].StyleIndex!.Value);

                CellFormat numberFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)numberStyleIndex);
                Assert.True(numberFormat.ApplyNumberFormat?.Value);
                NumberingFormat customFormat = stylesheet.NumberingFormats!.Elements<NumberingFormat>()
                    .First(format => format.NumberFormatId!.Value == numberFormat.NumberFormatId!.Value);
                Assert.Equal("0.00", customFormat.FormatCode!.Value);

                CellFormat fillFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)fillStyleIndex);
                Assert.True(fillFormat.ApplyFill?.Value);
                Fill fill = stylesheet.Fills!.Elements<Fill>().ElementAt((int)fillFormat.FillId!.Value);
                Assert.Equal("FFFFF2CC", fill.PatternFill!.ForegroundColor!.Rgb!.Value);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_GradientFills_AreReusableCellAndRangeStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.GradientFills.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellAt(1, 1).SetValue("Single").SetGradientFill("#FF0000", "#00FF00", 45);
                sheet.Range("B1:C1").SetGradientFill("112233", "445566", 90);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorkbookStylesPart stylesPart = spreadsheet.WorkbookPart!.WorkbookStylesPart!;
                Stylesheet stylesheet = stylesPart.Stylesheet!;
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value != null)
                    .ToDictionary(cell => cell.CellReference!.Value!);

                Fill singleFill = GetCellFill(stylesheet, cells["A1"]);
                AssertGradientFill(singleFill, "FFFF0000", "FF00FF00", 45D);

                uint rangeStyleIndex = cells["B1"].StyleIndex!.Value;
                Assert.Equal(rangeStyleIndex, cells["C1"].StyleIndex!.Value);
                Fill rangeFill = GetCellFill(stylesheet, cells["B1"]);
                AssertGradientFill(rangeFill, "FF112233", "FF445566", 90D);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }

            static Fill GetCellFill(Stylesheet stylesheet, Cell cell) {
                CellFormat format = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
                Assert.True(format.ApplyFill?.Value);
                return stylesheet.Fills!.Elements<Fill>().ElementAt((int)format.FillId!.Value);
            }

            static void AssertGradientFill(Fill fill, string expectedStart, string expectedEnd, double expectedDegree) {
                GradientFill gradient = fill.GradientFill!;
                Assert.NotNull(gradient);
                Assert.Equal(GradientValues.Linear, gradient.Type!.Value);
                Assert.Equal(expectedDegree, gradient.Degree!.Value);

                List<GradientStop> stops = gradient.Elements<GradientStop>().ToList();
                Assert.Equal(2, stops.Count);
                Assert.Equal(0D, stops[0].Position!.Value);
                Assert.Equal(expectedStart, stops[0].Color!.Rgb!.Value);
                Assert.Equal(1D, stops[1].Position!.Value);
                Assert.Equal(expectedEnd, stops[1].Color!.Rgb!.Value);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_HeaderColumnRanges_TargetWorksheetAndTableDataColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.HeaderColumnRanges.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Sales Amount");
                sheet.CellValue(2, 1, "NA");
                sheet.CellValue(2, 2, 100);
                sheet.CellValue(3, 1, "EMEA");
                sheet.CellValue(3, 2, 200);
                sheet.AddTable("A1:B3", hasHeader: true, name: "SalesTable", style: ExcelTableStyle.TableStyleMedium2);

                Assert.Equal("B2:B3", sheet.GetColumnRangeByHeader("Sales Amount"));
                Assert.Equal("B1:B3", sheet.GetColumnRangeByHeader("Sales Amount", includeHeader: true));
                Assert.Equal("B2:B3", sheet.GetColumnRangeByHeader("Sales   Amount", tableName: "SalesTable"));

                ExcelRange tableColumn = sheet.Table("SalesTable").Column("sales amount");
                Assert.Equal("B2:B3", tableColumn.Address);

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
                Assert.Equal("B2:B3", document["Data"].GetColumnRangeByHeader("Sales Amount", tableName: "SalesTable"));
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculationPolicy_EvaluatesSupportedFormulasBeforeSave() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.Calculation.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, 2d);
                sheet.CellValue(2, 1, 3d);
                sheet.CellFormula(3, 1, "SUM(A1:A2)");
                sheet.CellFormula(4, 1, "A1+A2");
                document.Calculation.EvaluateFormulasBeforeSave = true;
                document.Calculation.ForceFullCalculationOnOpen = true;
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Cell sumCell = worksheetPart.Worksheet.Descendants<Cell>().First(cell => cell.CellReference?.Value == "A3");
                Cell binaryCell = worksheetPart.Worksheet.Descendants<Cell>().First(cell => cell.CellReference?.Value == "A4");

                Assert.Equal("SUM(A1:A2)", sumCell.CellFormula!.Text);
                Assert.Equal("5", sumCell.CellValue!.Text);
                Assert.Equal("5", binaryCell.CellValue!.Text);
                Assert.True(spreadsheet.WorkbookPart.Workbook.CalculationProperties!.ForceFullCalculation!.Value);
                Assert.True(spreadsheet.WorkbookPart.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculateFacade_EvaluatesSupportedFormulaCaches() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculateFacade.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, 2d);
                sheet.CellValue(2, 1, 3d);
                sheet.CellFormula(3, 1, "SUM(A1:A2)");
                sheet.CellFormula(4, 1, "A1+A2");

                Assert.Equal(2, document.Calculate());

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(0, inspection.MissingCachedResults);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A3" && formula.CachedValue == "5");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A4" && formula.CachedValue == "5");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Cell sumCell = worksheetPart.Worksheet.Descendants<Cell>().First(cell => cell.CellReference?.Value == "A3");
                Cell binaryCell = worksheetPart.Worksheet.Descendants<Cell>().First(cell => cell.CellReference?.Value == "A4");

                Assert.Equal("5", sumCell.CellValue!.Text);
                Assert.Equal("5", binaryCell.CellValue!.Text);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculateFacade_EvaluatesSameSheetFormulaDependencies() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculateDependencies.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellFormula(1, 1, "A2+1");
                sheet.CellFormula(2, 1, "A3+1");
                sheet.CellFormula(3, 1, "SUM(B1:B2)");
                sheet.CellValue(1, 2, 2d);
                sheet.CellValue(2, 2, 3d);

                Assert.Equal(3, document.Calculate());

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(0, inspection.MissingCachedResults);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A1" && formula.CachedValue == "7");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A2" && formula.CachedValue == "6");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A3" && formula.CachedValue == "5");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value != null)
                    .ToDictionary(cell => cell.CellReference!.Value!);

                Assert.Equal("7", cells["A1"].CellValue!.Text);
                Assert.Equal("6", cells["A2"].CellValue!.Text);
                Assert.Equal("5", cells["A3"].CellValue!.Text);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculateFacade_EvaluatesCrossSheetNumericReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculateCrossSheetReferences.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet calc = document.AddWorksheet("Calc");
                ExcelSheet data = document.AddWorksheet("Data Sheet");
                ExcelSheet budget = document.AddWorksheet("Budget $ FY26");
                data.CellValue(1, 1, 2d);
                data.CellValue(2, 1, 3d);
                data.CellValue(1, 2, 4d);
                data.CellValue(2, 2, 6d);
                budget.CellValue(1, 1, 7d);

                calc.CellFormula(1, 1, "'Data Sheet'!A1+'Data Sheet'!A2");
                calc.CellFormula(2, 1, "SUM('Data Sheet'!B1:B2)");
                calc.CellFormula(3, 1, "IF('Data Sheet'!B2>5,10,0)");
                calc.CellFormula(4, 1, "'Budget $ FY26'!$A$1+1");
                calc.CellFormula(5, 1, "IF('Budget $ FY26'!$A$1>6,11,0)");
                calc.CellFormula(6, 1, "'Budget $ FY26'!A1+'Data Sheet'!A1");

                Assert.Equal(6, document.Calculate());

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(0, inspection.MissingCachedResults);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A1" && formula.CachedValue == "5");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A2" && formula.CachedValue == "10");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A3" && formula.CachedValue == "10");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A4" && formula.CachedValue == "8");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A5" && formula.CachedValue == "11");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A6" && formula.CachedValue == "9");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart calcPart = spreadsheet.WorkbookPart!.WorksheetParts
                    .First(part => part.Worksheet.Descendants<Cell>().Any(cell => cell.CellReference?.Value == "A3" && cell.CellFormula != null));
                Dictionary<string, Cell> cells = calcPart.Worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value != null)
                    .ToDictionary(cell => cell.CellReference!.Value!);

                Assert.Equal("5", cells["A1"].CellValue!.Text);
                Assert.Equal("10", cells["A2"].CellValue!.Text);
                Assert.Equal("10", cells["A3"].CellValue!.Text);
                Assert.Equal("8", cells["A4"].CellValue!.Text);
                Assert.Equal("11", cells["A5"].CellValue!.Text);
                Assert.Equal("9", cells["A6"].CellValue!.Text);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculateFacade_EvaluatesCrossSheetFormulaDependencies() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculateCrossSheetFormulaDependencies.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet calc = document.AddWorksheet("Calc");
                ExcelSheet data = document.AddWorksheet("Data");
                calc.CellFormula(1, 1, "Data!A1+1");
                data.CellFormula(1, 1, "A2+1");
                data.CellValue(2, 1, 4d);

                Assert.Equal(2, document.Calculate());

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(0, inspection.MissingCachedResults);
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Calc" && formula.CellReference == "A1" && formula.CachedValue == "6");
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Data" && formula.CellReference == "A1" && formula.CachedValue == "5");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart calcPart = spreadsheet.WorkbookPart!.WorksheetParts
                    .First(part => part.Worksheet.Descendants<Cell>().Any(cell => cell.CellReference?.Value == "A1" && cell.CellFormula?.Text == "Data!A1+1"));
                Cell calcCell = calcPart.Worksheet.Descendants<Cell>().First(cell => cell.CellReference?.Value == "A1");

                Assert.Equal("6", calcCell.CellValue!.Text);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculateFacade_EvaluatesNamedRangeReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculateNamedRanges.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet calc = document.AddWorksheet("Calc");
                ExcelSheet data = document.AddWorksheet("Data Sheet");
                data.CellValue(1, 1, 2d);
                data.CellValue(2, 1, 3d);
                data.CellValue(1, 2, 4d);

                document.SetNamedRange("GlobalValues", "'Data Sheet'!A1:A2", save: false);
                document.SetNamedRange("SharedInput", "'Data Sheet'!A1", save: false);
                data.SetNamedRange("SharedInput", "B1", save: false);

                calc.CellFormula(1, 1, "SUM(GlobalValues)");
                calc.CellFormula(2, 1, "SharedInput*10");
                calc.CellFormula(3, 1, "'Data Sheet'!SharedInput*5");
                data.CellFormula(3, 2, "SharedInput*10");

                Assert.Equal(4, document.Calculate());

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(0, inspection.MissingCachedResults);
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Calc" && formula.CellReference == "A1" && formula.CachedValue == "5");
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Calc" && formula.CellReference == "A2" && formula.CachedValue == "20");
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Calc" && formula.CellReference == "A3" && formula.CachedValue == "20");
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Data Sheet" && formula.CellReference == "B3" && formula.CachedValue == "40");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Calc" && formula.CellReference == "A1" && formula.CachedValue == "5");
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Calc" && formula.CellReference == "A2" && formula.CachedValue == "20");
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Calc" && formula.CellReference == "A3" && formula.CachedValue == "20");
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Data Sheet" && formula.CellReference == "B3" && formula.CachedValue == "40");
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculateFacade_EvaluatesTableReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculateTableReferences.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet calc = document.AddWorksheet("Calc");
                ExcelSheet data = document.AddWorksheet("Data");
                data.CellValue(1, 1, "Region");
                data.CellValue(1, 2, "Amount");
                data.CellValue(1, 3, "Tax");
                data.CellValue(2, 1, "EU");
                data.CellValue(2, 2, 10d);
                data.CellValue(2, 3, 1d);
                data.CellValue(3, 1, "EU");
                data.CellValue(3, 2, 20d);
                data.CellValue(3, 3, 2d);
                data.CellValue(4, 1, "US");
                data.CellValue(4, 2, 30d);
                data.CellValue(4, 3, 3d);
                data.Range("A1:C4").CreateTable("SalesData");

                calc.CellFormula(1, 1, "SUM(SalesData[Amount])");
                calc.CellFormula(2, 1, "SUM(SalesData[[#Data],[Tax]])");
                calc.CellFormula(3, 1, "COUNTIF(SalesData[Region],\"EU\")");
                calc.CellFormula(4, 1, "SUM(SalesData)");
                calc.CellFormula(5, 1, "XLOOKUP(\"US\",SalesData[Region],SalesData[Amount])");

                Assert.Equal(5, document.Calculate());

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(0, inspection.MissingCachedResults);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A1" && formula.CachedValue == "60");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A2" && formula.CachedValue == "6");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A3" && formula.CachedValue == "2");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A4" && formula.CachedValue == "66");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A5" && formula.CachedValue == "30");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A1" && formula.CachedValue == "60");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A2" && formula.CachedValue == "6");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A3" && formula.CachedValue == "2");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A4" && formula.CachedValue == "66");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A5" && formula.CachedValue == "30");
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculateFacade_EvaluatesCrossSheetTableFormulaDependencies() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculateTableFormulaDependencies.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet calc = document.AddWorksheet("Calc");
                ExcelSheet data = document.AddWorksheet("Data");
                data.CellValue(1, 1, "Region");
                data.CellValue(1, 2, "Base");
                data.CellValue(1, 3, "Amount");
                data.CellValue(2, 1, "EU");
                data.CellValue(2, 2, 10d);
                data.CellFormula(2, 3, "B2*2");
                data.CellValue(3, 1, "US");
                data.CellValue(3, 2, 5d);
                data.CellFormula(3, 3, "B3*3");
                data.Range("A1:C3").CreateTable("SalesData");

                calc.CellFormula(1, 1, "SUM(SalesData[Amount])");

                Assert.Equal(3, document.Calculate());

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(0, inspection.MissingCachedResults);
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Calc" && formula.CellReference == "A1" && formula.CachedValue == "35");
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Data" && formula.CellReference == "C2" && formula.CachedValue == "20");
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Data" && formula.CellReference == "C3" && formula.CachedValue == "15");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Contains(inspection.Formulas, formula => formula.SheetName == "Calc" && formula.CellReference == "A1" && formula.CachedValue == "35");
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculateFacade_EvaluatesTextHelpersAndLookupReturns() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculateTextHelpers.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Text");
                sheet.CellValue(1, 1, "North");
                sheet.CellValue(2, 1, "South");
                sheet.CellValue(3, 1, "  East   Hub  ");
                sheet.CellValue(5, 1, "EU");
                sheet.CellValue(5, 2, "Europe");
                sheet.CellValue(6, 1, "US");
                sheet.CellValue(6, 2, "United States");

                sheet.CellFormula(1, 3, "CONCAT(A1,\"-\",A2)");
                sheet.CellFormula(2, 3, "TEXTJOIN(\",\",TRUE,A1:A3)");
                sheet.CellFormula(3, 3, "LEFT(A2,2)");
                sheet.CellFormula(4, 3, "RIGHT(A2,3)");
                sheet.CellFormula(5, 3, "MID(A2,2,3)");
                sheet.CellFormula(6, 3, "LEN(TRIM(A3))");
                sheet.CellFormula(7, 3, "TRIM(A3)");
                sheet.CellFormula(8, 3, "XLOOKUP(\"US\",A5:A6,B5:B6)");
                sheet.CellFormula(9, 3, "VLOOKUP(\"EU\",A5:B6,2,FALSE)");
                sheet.CellFormula(10, 3, "CONCAT(\"A\"\"B\",\"-\",A1)");

                Assert.Equal(10, document.Calculate());

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(0, inspection.MissingCachedResults);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "C1" && formula.CachedValue == "North-South");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "C2" && formula.CachedValue == "North,South,  East   Hub  ");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "C3" && formula.CachedValue == "So");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "C4" && formula.CachedValue == "uth");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "C5" && formula.CachedValue == "out");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "C6" && formula.CachedValue == "8");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "C7" && formula.CachedValue == "East Hub");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "C8" && formula.CachedValue == "United States");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "C9" && formula.CachedValue == "Europe");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "C10" && formula.CachedValue == "A\"B-North");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value != null)
                    .ToDictionary(cell => cell.CellReference!.Value!);

                Assert.Equal(CellValues.String, cells["C1"].DataType!.Value);
                Assert.Equal("North-South", cells["C1"].CellValue!.Text);
                Assert.Equal(CellValues.Number, cells["C6"].DataType!.Value);
                Assert.Equal("8", cells["C6"].CellValue!.Text);
                Assert.Equal("United States", cells["C8"].CellValue!.Text);
                Assert.Equal("Europe", cells["C9"].CellValue!.Text);
                Assert.Equal("A\"B-North", cells["C10"].CellValue!.Text);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculateFacade_LeavesCircularFormulaDependenciesUncached() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculateCircularDependencies.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellFormula(1, 1, "A2+1");
                sheet.CellFormula(2, 1, "A1+1");

                Assert.Equal(0, document.Calculate());

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(2, inspection.MissingCachedResults);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A1" && formula.CachedValue == null);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "A2" && formula.CachedValue == null);
                document.Save();
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculateFacade_LeavesTextFormulasWithUnresolvedDependenciesUncached() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculateUnresolvedTextDependencies.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellFormula(1, 1, "A2+1");
                sheet.CellFormula(2, 1, "A1+1");
                sheet.CellFormula(1, 2, "CONCAT(A1,\"x\")");
                sheet.CellFormula(2, 2, "TEXTJOIN(\",\",TRUE,A1:A2)");

                Assert.Equal(0, document.Calculate());

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(4, inspection.MissingCachedResults);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "B1" && formula.CachedValue == null);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "B2" && formula.CachedValue == null);
                document.Save();
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_SaveOptions_CanEvaluateFormulasForSingleSave() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculationSaveOptions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellFormula(1, 1, "A3*2");
                sheet.CellValue(1, 2, 2d);
                sheet.CellValue(2, 2, 3d);
                sheet.CellFormula(3, 1, "SUM(B1:B2)");

                document.Save(filePath, new ExcelSaveOptions {
                    EvaluateFormulasBeforeSave = true,
                    ForceFullCalculationOnOpen = true
                });

                Assert.False(document.Calculation.EvaluateFormulasBeforeSave);
                Assert.False(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Cell dependentCell = worksheetPart.Worksheet.Descendants<Cell>().First(cell => cell.CellReference?.Value == "A1");
                Cell sumCell = worksheetPart.Worksheet.Descendants<Cell>().First(cell => cell.CellReference?.Value == "A3");

                Assert.Equal("10", dependentCell.CellValue!.Text);
                Assert.Equal("5", sumCell.CellValue!.Text);
                Assert.True(spreadsheet.WorkbookPart.Workbook.CalculationProperties!.ForceFullCalculation!.Value);
                Assert.True(spreadsheet.WorkbookPart.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_FormulaInspection_ReportsSpecificUnsupportedReasons() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.FormulaUnsupportedReasons.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Diagnostics");
                sheet.CellValue(1, 1, "North");
                sheet.CellValue(2, 1, "South");

                sheet.CellFormula(1, 2, "UNIQUE(A1:A2)");
                sheet.CellFormula(2, 2, "SUM(A1;A2)");
                sheet.CellFormula(3, 2, "A1&A2");
                sheet.CellFormula(4, 2, "SUM({1,2})");

                ExcelFormulaInspection inspection = document.InspectFormulas();
                Assert.Equal(4, inspection.UnsupportedFormulas);
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "B1"
                    && formula.UnsupportedReason == "Function 'UNIQUE' is not supported by OfficeIMO's lightweight evaluator.");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "B2"
                    && formula.UnsupportedReason == "Formula uses semicolon argument separators; OfficeIMO's lightweight evaluator expects Open XML comma-separated formulas.");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "B3"
                    && formula.UnsupportedReason == "Formula uses the text concatenation operator, which OfficeIMO's lightweight evaluator does not currently support.");
                Assert.Contains(inspection.Formulas, formula => formula.CellReference == "B4"
                    && formula.UnsupportedReason == "Formula uses array constants, which OfficeIMO's lightweight evaluator does not currently support.");
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_FormulaInspection_ReportsSupportAndCacheStatus() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.FormulaInspection.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, 2d);
                sheet.CellValue(2, 1, 3d);
                sheet.CellValue(1, 2, "EU");
                sheet.CellValue(2, 2, "US");
                sheet.CellValue(3, 2, "EU");
                sheet.CellValue(4, 2, "EMEA");
                sheet.CellValue(1, 3, 10d);
                sheet.CellValue(2, 3, 20d);
                sheet.CellValue(3, 3, 30d);
                sheet.CellValue(4, 3, 40d);
                sheet.CellValue(1, 4, new DateTime(2024, 5, 21));
                sheet.CellValue(2, 4, new DateTime(1899, 12, 30, 14, 30, 45).ToOADate());
                sheet.CellValue(3, 4, new DateTime(2024, 5, 27));
                sheet.CellValue(1, 5, "EU");
                sheet.CellValue(1, 6, "US");
                sheet.CellValue(1, 7, "EMEA");
                sheet.CellValue(2, 5, 100d);
                sheet.CellValue(2, 6, 200d);
                sheet.CellValue(2, 7, 300d);
                sheet.CellFormula(3, 1, "SUM(A1:A2)");
                sheet.CellFormula(4, 1, "VLOOKUP(A1,B1:C2,2,FALSE)");
                sheet.CellFormula(5, 1, "ABS(-4)");
                sheet.CellFormula(6, 1, "ROUND(2.345,2)");
                sheet.CellFormula(7, 1, "A1+A2");
                sheet.CellFormula(8, 1, "ROUNDUP(2.341,2)");
                sheet.CellFormula(9, 1, "ROUNDDOWN(2.349,2)");
                sheet.CellFormula(10, 1, "INT(2.9)");
                sheet.CellFormula(11, 1, "POWER(2,3)");
                sheet.CellFormula(12, 1, "SQRT(9)");
                sheet.CellFormula(13, 1, "MOD(10,3)");
                sheet.CellFormula(14, 1, "IF(A2>A1,10,0)");
                sheet.CellFormula(15, 1, "AND(A1>0,A2>=3)");
                sheet.CellFormula(16, 1, "OR(A1>10,A2=3)");
                sheet.CellFormula(17, 1, "IFERROR(SUM(A1:A2),0)");
                sheet.CellFormula(18, 1, "IFERROR(A1/0,99)");
                sheet.CellFormula(19, 1, "IF(A2>A1,SUM(A1:A2),0)");
                sheet.CellFormula(20, 1, "IFERROR(A1/0,SUM(A1:A2))");
                sheet.CellFormula(21, 1, "IF(AND(A1>0,A2>=3),20,0)");
                sheet.CellFormula(22, 1, "IF(OR(A1>10,A2=3),30,0)");
                sheet.CellFormula(23, 1, "NOT(A1>10)");
                sheet.CellFormula(24, 1, "IF(NOT(A1>10),40,0)");
                sheet.CellFormula(25, 1, "COUNTIF(B1:B4,\"EU\")");
                sheet.CellFormula(26, 1, "COUNTIF(C1:C4,\">=20\")");
                sheet.CellFormula(27, 1, "SUMIF(B1:B4,\"EU\",C1:C4)");
                sheet.CellFormula(28, 1, "AVERAGEIF(C1:C4,\">20\",C1:C4)");
                sheet.CellFormula(29, 1, "COUNTIF(B1:B4,\"E*\")");
                sheet.CellFormula(30, 1, "COUNTIFS(B1:B4,\"EU\",C1:C4,\">=20\")");
                sheet.CellFormula(31, 1, "SUMIFS(C1:C4,B1:B4,\"E*\",C1:C4,\">=30\")");
                sheet.CellFormula(32, 1, "AVERAGEIFS(C1:C4,B1:B4,\"<>US\",C1:C4,\">=30\")");
                sheet.CellFormula(33, 1, "DATE(2024,5,21)");
                sheet.CellFormula(34, 1, "YEAR(D1)");
                sheet.CellFormula(35, 1, "MONTH(D1)");
                sheet.CellFormula(36, 1, "DAY(D1)");
                sheet.CellFormula(37, 1, "TIME(14,30,45)");
                sheet.CellFormula(38, 1, "HOUR(D2)");
                sheet.CellFormula(39, 1, "MINUTE(D2)");
                sheet.CellFormula(40, 1, "SECOND(D2)");
                sheet.CellFormula(41, 1, "TODAY()");
                sheet.CellFormula(42, 1, "NOW()");
                sheet.CellFormula(43, 1, "EDATE(D1,1)");
                sheet.CellFormula(44, 1, "EOMONTH(D1,0)");
                sheet.CellFormula(45, 1, "DAYS(DATE(2024,5,31),D1)");
                sheet.CellFormula(46, 1, "WEEKDAY(D1,2)");
                sheet.CellFormula(47, 1, "NETWORKDAYS(D1,DATE(2024,5,31),D3:D3)");
                sheet.CellFormula(48, 1, "PRODUCT(A1:A2,2)");
                sheet.CellFormula(49, 1, "MEDIAN(C1:C4)");
                sheet.CellFormula(50, 1, "LARGE(C1:C4,2)");
                sheet.CellFormula(51, 1, "SMALL(C1:C4,3)");
                sheet.CellFormula(52, 1, "SUMPRODUCT(A1:A2,C1:C2)");
                sheet.CellFormula(53, 1, "SIGN(-10)");
                sheet.CellFormula(54, 1, "TRUNC(2.987,2)");
                sheet.CellFormula(55, 1, "CEILING(2.1,0.5)");
                sheet.CellFormula(56, 1, "FLOOR(2.9,0.5)");
                sheet.CellFormula(57, 1, "LN(EXP(1))");
                sheet.CellFormula(58, 1, "LOG10(100)");
                sheet.CellFormula(59, 1, "EXP(1)");
                sheet.CellFormula(60, 1, "PI()");
                sheet.CellFormula(61, 1, "RADIANS(180)");
                sheet.CellFormula(62, 1, "DEGREES(PI())");
                sheet.CellFormula(63, 1, "VLOOKUP(\"US\",B1:C4,2,FALSE)");
                sheet.CellFormula(64, 1, "HLOOKUP(\"US\",E1:G2,2,FALSE)");
                sheet.CellFormula(65, 1, "XLOOKUP(\"EMEA\",B1:B4,C1:C4)");
                sheet.CellFormula(66, 1, "MINIFS(C1:C4,B1:B4,\"EU\")");
                sheet.CellFormula(67, 1, "MAXIFS(C1:C4,B1:B4,\"E*\",C1:C4,\">=30\")");
                sheet.CellFormula(68, 1, "COUNTBLANK(H1:H3)");
                sheet.CellFormula(69, 1, "SUBTOTAL(9,C1:C4)");

                ExcelFormulaInspection sheetInspection = sheet.InspectFormulas();
                Assert.Equal(67, sheetInspection.TotalFormulas);
                Assert.Equal(66, sheetInspection.SupportedFormulas);
                Assert.Equal(1, sheetInspection.UnsupportedFormulas);
                Assert.Equal(67, sheetInspection.MissingCachedResults);
                Assert.False(sheetInspection.AllSupported);
                Assert.False(sheetInspection.AllHaveCachedResults);
                Assert.Contains("SUM", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("ABS", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("SIGN", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("ROUND", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("ROUNDUP", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("ROUNDDOWN", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("TRUNC", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("COUNTIF", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("SUMIF", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("AVERAGEIF", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("COUNTIFS", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("SUMIFS", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("AVERAGEIFS", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("MINIFS", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("MAXIFS", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("COUNTBLANK", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("SUBTOTAL", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("PRODUCT", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("MEDIAN", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("LARGE", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("SMALL", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("SUMPRODUCT", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("VLOOKUP", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("HLOOKUP", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("XLOOKUP", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("CONCAT", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("TEXTJOIN", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("LEFT", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("RIGHT", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("MID", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("LEN", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("TRIM", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("DATE", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("TIME", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("TODAY", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("NOW", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("YEAR", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("MONTH", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("DAY", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("HOUR", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("MINUTE", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("SECOND", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("EDATE", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("EOMONTH", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("DAYS", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("WEEKDAY", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("NETWORKDAYS", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("INT", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("CEILING", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("FLOOR", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("POWER", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("SQRT", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("LN", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("LOG10", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("EXP", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("PI", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("RADIANS", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("DEGREES", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("MOD", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("IF", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("AND", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("OR", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("NOT", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("IFERROR", sheetInspection.Capabilities.SupportedFunctions);
                Assert.Contains("+", sheetInspection.Capabilities.SupportedOperators);
                Assert.Contains(">=", sheetInspection.Capabilities.SupportedOperators);
                Assert.Contains("same-sheet A1 range", string.Join(";", sheetInspection.Capabilities.SupportedOperandKinds));
                Assert.Contains("same-sheet numeric/text comparison", string.Join(";", sheetInspection.Capabilities.SupportedOperandKinds));
                Assert.Equal(8192, sheetInspection.Capabilities.MaxFormulaLength);
                Assert.Contains(sheetInspection.Formulas, formula => formula.CellReference == "A3" && formula.IsSupportedByOfficeIMO);
                Assert.Contains(sheetInspection.Formulas, formula => formula.CellReference == "A4"
                    && !formula.IsSupportedByOfficeIMO
                    && !string.IsNullOrWhiteSpace(formula.UnsupportedReason));
                Assert.Contains("| Calc | A4 | VLOOKUP(A1,B1:C2,2,FALSE) | no | no | no |", sheetInspection.ToMarkdown());
                InvalidOperationException unsupportedException = Assert.Throws<InvalidOperationException>(() => sheetInspection.EnsureAllSupported());
                Assert.Contains("Calc!A4", unsupportedException.Message);
                InvalidOperationException missingCacheException = Assert.Throws<InvalidOperationException>(() => sheetInspection.EnsureAllHaveCachedResults());
                Assert.Contains("Calc!A3", missingCacheException.Message);

                ExcelFormulaInspection workbookInspection = document.InspectFormulas();
                Assert.Equal(sheetInspection.TotalFormulas, workbookInspection.TotalFormulas);
                Assert.Equal("Calc", workbookInspection.Formulas[0].SheetName);

                double nowBefore = DateTime.Now.ToOADate();
                Assert.Equal(66, document.RecalculateSupportedFormulas());
                ExcelFormulaInspection afterRecalculate = document.InspectFormulas();
                double nowAfter = DateTime.Now.ToOADate();
                Assert.Equal(1, afterRecalculate.MissingCachedResults);
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A3" && formula.CachedValue == "5");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A5" && formula.CachedValue == "4");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A6" && formula.CachedValue == "2.35");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A7" && formula.CachedValue == "5");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A8" && formula.CachedValue == "2.35");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A9" && formula.CachedValue == "2.34");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A10" && formula.CachedValue == "2");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A11" && formula.CachedValue == "8");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A12" && formula.CachedValue == "3");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A13" && formula.CachedValue == "1");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A14" && formula.CachedValue == "10");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A15" && formula.CachedValue == "1");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A16" && formula.CachedValue == "1");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A17" && formula.CachedValue == "5");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A18" && formula.CachedValue == "99");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A19" && formula.CachedValue == "5");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A20" && formula.CachedValue == "5");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A21" && formula.CachedValue == "20");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A22" && formula.CachedValue == "30");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A23" && formula.CachedValue == "1");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A24" && formula.CachedValue == "40");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A25" && formula.CachedValue == "2");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A26" && formula.CachedValue == "3");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A27" && formula.CachedValue == "40");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A28" && formula.CachedValue == "35");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A29" && formula.CachedValue == "3");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A30" && formula.CachedValue == "1");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A31" && formula.CachedValue == "70");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A32" && formula.CachedValue == "35");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A33" && formula.CachedValue == new DateTime(2024, 5, 21).ToOADate().ToString(CultureInfo.InvariantCulture));
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A34" && formula.CachedValue == "2024");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A35" && formula.CachedValue == "5");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A36" && formula.CachedValue == "21");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A37" && formula.CachedValue == (52245d / 86400d).ToString(CultureInfo.InvariantCulture));
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A38" && formula.CachedValue == "14");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A39" && formula.CachedValue == "30");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A40" && formula.CachedValue == "45");
                string todayCached = Assert.Single(afterRecalculate.Formulas, formula => formula.CellReference == "A41").CachedValue!;
                Assert.Equal(DateTime.Today, DateTime.FromOADate(double.Parse(todayCached, CultureInfo.InvariantCulture)).Date);
                string nowCached = Assert.Single(afterRecalculate.Formulas, formula => formula.CellReference == "A42").CachedValue!;
                double nowCachedValue = double.Parse(nowCached, CultureInfo.InvariantCulture);
                Assert.InRange(nowCachedValue, nowBefore, nowAfter);
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A43" && formula.CachedValue == new DateTime(2024, 6, 21).ToOADate().ToString(CultureInfo.InvariantCulture));
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A44" && formula.CachedValue == new DateTime(2024, 5, 31).ToOADate().ToString(CultureInfo.InvariantCulture));
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A45" && formula.CachedValue == "10");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A46" && formula.CachedValue == "2");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A47" && formula.CachedValue == "8");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A48" && formula.CachedValue == "12");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A49" && formula.CachedValue == "25");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A50" && formula.CachedValue == "30");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A51" && formula.CachedValue == "30");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A52" && formula.CachedValue == "80");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A53" && formula.CachedValue == "-1");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A54" && formula.CachedValue == "2.98");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A55" && formula.CachedValue == "2.5");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A56" && formula.CachedValue == "2.5");
                Assert.InRange(double.Parse(Assert.Single(afterRecalculate.Formulas, formula => formula.CellReference == "A57").CachedValue!, CultureInfo.InvariantCulture), 0.999999999d, 1.000000001d);
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A58" && formula.CachedValue == "2");
                Assert.InRange(double.Parse(Assert.Single(afterRecalculate.Formulas, formula => formula.CellReference == "A59").CachedValue!, CultureInfo.InvariantCulture), Math.E - 0.000000001d, Math.E + 0.000000001d);
                Assert.InRange(double.Parse(Assert.Single(afterRecalculate.Formulas, formula => formula.CellReference == "A60").CachedValue!, CultureInfo.InvariantCulture), Math.PI - 0.000000001d, Math.PI + 0.000000001d);
                Assert.InRange(double.Parse(Assert.Single(afterRecalculate.Formulas, formula => formula.CellReference == "A61").CachedValue!, CultureInfo.InvariantCulture), Math.PI - 0.000000001d, Math.PI + 0.000000001d);
                Assert.InRange(double.Parse(Assert.Single(afterRecalculate.Formulas, formula => formula.CellReference == "A62").CachedValue!, CultureInfo.InvariantCulture), 179.999999999d, 180.000000001d);
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A63" && formula.CachedValue == "20");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A64" && formula.CachedValue == "200");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A65" && formula.CachedValue == "40");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A66" && formula.CachedValue == "10");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A67" && formula.CachedValue == "40");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A68" && formula.CachedValue == "3");
                Assert.Contains(afterRecalculate.Formulas, formula => formula.CellReference == "A69" && formula.CachedValue == "100");
                Assert.Throws<InvalidOperationException>(() => afterRecalculate.EnsureAllSupported());
                Assert.Throws<InvalidOperationException>(() => afterRecalculate.EnsureAllHaveCachedResults());

                document.InvalidateFormulas();
                Assert.Equal(67, document.InspectFormulas().DirtyFormulas);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_FeatureReport_SummarizesEditablePartialAndPreservedFeatures() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.FeatureReport.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Sales");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, "EU");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10d);
                sheet.CellValue(3, 1, "EU");
                sheet.CellValue(3, 2, "B");
                sheet.CellValue(3, 3, 20d);
                sheet.CellValue(4, 1, "US");
                sheet.CellValue(4, 2, "A");
                sheet.CellValue(4, 3, 30d);
                sheet.CellFormula(5, 3, "SUM(C2:C4)");

                sheet.Range("A1:C4").CreateTable("SalesData");
                sheet.Range("B2:B4").Validate.List("A", "B");
                sheet.Range("C2:C4").ConditionalFormat.DataBar("#5B9BD5");
                sheet.Sparklines("C2:C4").Column().At("D2:D4");
                sheet.Pivot("A1:C4").Rows("Region").Columns("Product").Sum("Sales").At("F2", "SalesPivot");
                sheet.Chart("A1:C4").ColumnClustered().Title("Sales").At(12, 1);

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                ExcelFeatureFinding worksheets = Assert.Single(report.Features, feature => feature.Name == "Worksheets");
                Assert.Equal(ExcelFeatureSupportLevel.Editable, worksheets.SupportLevel);

                Assert.Contains(report.EditableFeatures, feature => feature.Name == "Tables" && feature.Count == 1);
                Assert.Contains(report.EditableFeatures, feature => feature.Name == "Data validations" && feature.Count == 1);
                Assert.Contains(report.EditableFeatures, feature => feature.Name == "Sparklines" && feature.Count == 3);
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Charts" && feature.Count == 1);
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Pivot tables" && feature.Count == 1);
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Missing formula caches" && feature.Count == 1);
                Assert.True(report.HasAdvancedFeatures);
                Assert.Same(report, report.EnsureNoUnsupportedFeatures());
                InvalidOperationException advancedException = Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                Assert.Contains("Missing formula caches", advancedException.Message);
                string markdown = report.ToMarkdown();
                Assert.Contains("# Excel Feature Report", markdown);
                Assert.Contains("| Visualization | Pivot tables | 1 | PartiallyEditable |", markdown);
                Assert.Contains("| Calculation | Missing formula caches | 1 | Preserved |", markdown);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_FeatureReport_IncludesPreservationDetails() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.FeatureReportDetails.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Links");
                sheet.CellValue(1, 1, "Resource");
                sheet.CellValue(2, 1, "Spec");
                sheet.SetHyperlink(2, 1, "https://example.org/spec", display: "Spec");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                CustomXmlPart customXmlPart = spreadsheet.WorkbookPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using var stream = new MemoryStream(Encoding.UTF8.GetBytes("<metadata><owner>OfficeIMO</owner></metadata>"));
                customXmlPart.FeedData(stream);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                ExcelFeatureFinding externalLinks = Assert.Single(report.PartiallyEditableFeatures, feature => feature.Name == "External hyperlinks");
                Assert.Equal(1, externalLinks.Count);
                Assert.Contains(externalLinks.Details, detail => detail.Contains("https://example.org/spec", StringComparison.OrdinalIgnoreCase));

                ExcelFeatureFinding customXml = Assert.Single(report.PreservedFeatures, feature => feature.Name == "Custom XML parts");
                Assert.Equal(1, customXml.Count);
                Assert.Contains(customXml.Details, detail => detail.Contains("/customXml/", StringComparison.OrdinalIgnoreCase));

                string markdown = report.ToMarkdown();
                Assert.Contains("https://example.org/spec", markdown);
                Assert.Contains("/customXml/", markdown);
            }
        }

        [Fact]
        public void Test_ExcelFeatureReport_DetectsSlicerAndTimelinePackageParts() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFeatureReport.SlicerTimelineParts.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Dashboard");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(2, 1, "EU");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;

                ExtendedPart slicerPart = workbookPart.AddExtendedPart(
                    "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                    "application/vnd.ms-excel.slicerCache+xml",
                    "xml");
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" name=\"Region\"/>"))) {
                    slicerPart.FeedData(stream);
                }

                ExtendedPart timelinePart = workbookPart.AddExtendedPart(
                    "http://schemas.microsoft.com/office/2011/relationships/timelineCache",
                    "application/vnd.ms-excel.timelineCache+xml",
                    "xml");
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("<timelineCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2011/1/main\" name=\"OrderDate\"/>"))) {
                    timelinePart.FeedData(stream);
                }
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                ExcelFeatureFinding slicers = Assert.Single(report.FindFeatures("Slicers"));
                Assert.Equal(ExcelFeatureSupportLevel.PartiallyEditable, slicers.SupportLevel);
                Assert.Equal(1, slicers.Count);
                Assert.Contains(slicers.Details, detail => detail.Contains("slicerCache", StringComparison.OrdinalIgnoreCase));

                ExcelFeatureFinding timelines = Assert.Single(report.FindFeatures("Timelines"));
                Assert.Equal(ExcelFeatureSupportLevel.PartiallyEditable, timelines.SupportLevel);
                Assert.Equal(1, timelines.Count);
                Assert.Contains(timelines.Details, detail => detail.Contains("timelineCache", StringComparison.OrdinalIgnoreCase));

                Assert.Same(report, report.EnsureNoAdvancedFeatures());
                Assert.Same(report, report.EnsureCan(ExcelPreflightCapability.EditWorkbookStructure));
            }
        }

        [Fact]
        public void Test_ExcelFeatureReport_DetectsConnectionAndQueryTablePackageParts() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelFeatureReport.ConnectionQueryParts.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(2, 1, "EU");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();

                ExtendedPart connectionPart = workbookPart.AddExtendedPart(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
                    "xml");
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\"><connection id=\"1\" name=\"SalesConnection\" type=\"5\" refreshedVersion=\"7\"/></connections>"))) {
                    connectionPart.FeedData(stream);
                }

                ExtendedPart queryTablePart = worksheetPart.AddExtendedPart(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
                    "xml");
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("<queryTable xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"SalesQuery\" connectionId=\"1\"/>"))) {
                    queryTablePart.FeedData(stream);
                }
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                ExcelFeatureFinding connections = Assert.Single(report.FindFeatures("Connections and query tables"));
                Assert.Equal(ExcelFeatureSupportLevel.Preserved, connections.SupportLevel);
                Assert.Equal(2, connections.Count);
                Assert.Contains(connections.Details, detail => detail.Contains("connections", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(connections.Details, detail => detail.Contains("queryTable", StringComparison.OrdinalIgnoreCase));

                InvalidOperationException advancedException = Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                Assert.Contains("Connections and query tables", advancedException.Message);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_RoundTrip_PreservesSlicerAndTimelinePackageParts() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.PreserveSlicerTimelineParts.xlsx");
            byte[] slicerBytes = Encoding.UTF8.GetBytes("<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" name=\"Region\"/>");
            byte[] timelineBytes = Encoding.UTF8.GetBytes("<timelineCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2011/1/main\" name=\"OrderDate\"/>");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Dashboard");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(2, 1, "EU");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;

                ExtendedPart slicerPart = workbookPart.AddExtendedPart(
                    "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                    "application/vnd.ms-excel.slicerCache+xml",
                    "xml");
                using (var stream = new MemoryStream(slicerBytes)) {
                    slicerPart.FeedData(stream);
                }

                ExtendedPart timelinePart = workbookPart.AddExtendedPart(
                    "http://schemas.microsoft.com/office/2011/relationships/timelineCache",
                    "application/vnd.ms-excel.timelineCache+xml",
                    "xml");
                using (var stream = new MemoryStream(timelineBytes)) {
                    timelinePart.FeedData(stream);
                }
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document["Dashboard"].CellValue(3, 1, "US");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                OpenXmlPart slicerPart = Assert.Single(workbookPart.Parts.Select(part => part.OpenXmlPart),
                    part => part.ContentType.IndexOf("slicerCache", StringComparison.OrdinalIgnoreCase) >= 0);
                OpenXmlPart timelinePart = Assert.Single(workbookPart.Parts.Select(part => part.OpenXmlPart),
                    part => part.ContentType.IndexOf("timelineCache", StringComparison.OrdinalIgnoreCase) >= 0);

                using (var stream = slicerPart.GetStream(FileMode.Open, FileAccess.Read)) {
                    using var buffer = new MemoryStream();
                    stream.CopyTo(buffer);
                    Assert.Equal(slicerBytes, buffer.ToArray());
                }

                using (var stream = timelinePart.GetStream(FileMode.Open, FileAccess.Read)) {
                    using var buffer = new MemoryStream();
                    stream.CopyTo(buffer);
                    Assert.Equal(timelineBytes, buffer.ToArray());
                }
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Slicers"
                    && feature.Details.Any(detail => detail.Contains("slicerCache", StringComparison.OrdinalIgnoreCase)));
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Timelines"
                    && feature.Details.Any(detail => detail.Contains("timelineCache", StringComparison.OrdinalIgnoreCase)));
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_RoundTrip_PreservesConnectionAndQueryTablePackageParts() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.PreserveConnectionQueryParts.xlsx");
            byte[] connectionBytes = Encoding.UTF8.GetBytes("<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\"><connection id=\"1\" name=\"SalesConnection\" type=\"5\" refreshedVersion=\"7\"/></connections>");
            byte[] queryTableBytes = Encoding.UTF8.GetBytes("<queryTable xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"SalesQuery\" connectionId=\"1\"/>");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(2, 1, "EU");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();

                ExtendedPart connectionPart = workbookPart.AddExtendedPart(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
                    "xml");
                using (var stream = new MemoryStream(connectionBytes)) {
                    connectionPart.FeedData(stream);
                }

                ExtendedPart queryTablePart = worksheetPart.AddExtendedPart(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
                    "xml");
                using (var stream = new MemoryStream(queryTableBytes)) {
                    queryTablePart.FeedData(stream);
                }
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document["Data"].CellValue(3, 1, "US");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                OpenXmlPart connectionPart = Assert.Single(workbookPart.Parts.Select(part => part.OpenXmlPart),
                    part => part.ContentType.IndexOf("connections", StringComparison.OrdinalIgnoreCase) >= 0);
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
                OpenXmlPart queryTablePart = Assert.Single(worksheetPart.Parts.Select(part => part.OpenXmlPart),
                    part => part.ContentType.IndexOf("queryTable", StringComparison.OrdinalIgnoreCase) >= 0);

                using (var stream = connectionPart.GetStream(FileMode.Open, FileAccess.Read)) {
                    using var buffer = new MemoryStream();
                    stream.CopyTo(buffer);
                    Assert.Equal(connectionBytes, buffer.ToArray());
                }

                using (var stream = queryTablePart.GetStream(FileMode.Open, FileAccess.Read)) {
                    using var buffer = new MemoryStream();
                    stream.CopyTo(buffer);
                    Assert.Equal(queryTableBytes, buffer.ToArray());
                }
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Connections and query tables"
                    && feature.Count == 2
                    && feature.Details.Any(detail => detail.Contains("queryTable", StringComparison.OrdinalIgnoreCase)));
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_RoundTrip_PreservesThreadedCommentsAndReportsDetails() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.PreserveThreadedComments.xlsx");
            const string personId = "{11111111-1111-1111-1111-111111111111}";
            const string commentId = "{22222222-2222-2222-2222-222222222222}";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Review");
                sheet.CellValue(1, 1, "Revenue");
                sheet.CellValue(2, 1, 1250d);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                WorkbookPersonPart personPart = workbookPart.AddNewPart<WorkbookPersonPart>();
                personPart.PersonList = new Threaded.PersonList(
                    new Threaded.Person {
                        DisplayName = "Modern Reviewer",
                        Id = personId,
                        UserId = "modern.reviewer@example.test",
                        ProviderId = "OfficeIMO.Tests"
                    });
                personPart.PersonList.Save();

                WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
                WorksheetThreadedCommentsPart threadedPart = worksheetPart.AddNewPart<WorksheetThreadedCommentsPart>();
                threadedPart.ThreadedComments = new Threaded.ThreadedComments(
                    new Threaded.ThreadedComment(
                        new Threaded.ThreadedCommentText("Confirm revenue threshold"))
                    {
                        Ref = "A1",
                        PersonId = personId,
                        Id = commentId,
                        DT = new DateTime(2026, 5, 29, 12, 0, 0, DateTimeKind.Utc),
                        Done = false
                    });
                threadedPart.ThreadedComments.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document["Review"].CellValue(3, 1, "Checked");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                WorkbookPersonPart personPart = Assert.Single(workbookPart.WorkbookPersonParts);
                Assert.Equal("Modern Reviewer", Assert.Single(personPart.PersonList!.Elements<Threaded.Person>()).DisplayName!.Value);

                WorksheetThreadedCommentsPart threadedPart = Assert.Single(workbookPart.WorksheetParts.Single().WorksheetThreadedCommentsParts);
                Threaded.ThreadedComment threadedComment = Assert.Single(threadedPart.ThreadedComments!.Elements<Threaded.ThreadedComment>());
                Assert.Equal("A1", threadedComment.Ref!.Value);
                Assert.Equal(commentId, threadedComment.Id!.Value);
                Assert.Equal("Confirm revenue threshold", threadedComment.ThreadedCommentText!.InnerText);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();
                ExcelFeatureFinding threadedComments = Assert.Single(report.FindFeatures("Threaded comments"));
                Assert.Equal(ExcelFeatureSupportLevel.PartiallyEditable, threadedComments.SupportLevel);
                Assert.Equal(1, threadedComments.Count);
                Assert.Contains(threadedComments.Details, detail => detail.Contains("Review: A1 by Modern Reviewer", StringComparison.OrdinalIgnoreCase));

                Assert.Same(report, report.EnsureNoAdvancedFeatures());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_RoundTrip_PreservesOleObjectAndFormControlMarkers() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.PreserveOleFormControls.xlsx");
            const string oleObjectsXml = "<x:oleObjects xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><x:oleObject progId=\"Package\" shapeId=\"1025\" r:id=\"rIdOlePackage\" /></x:oleObjects>";
            const string controlsXml = "<x:controls xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><x:control shapeId=\"1026\" name=\"ApproveButton\" r:id=\"rIdControl1\" /></x:controls>";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Controls");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(2, 1, "Before");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
                worksheetPart.Worksheet.Append(new OleObjects(oleObjectsXml));
                worksheetPart.Worksheet.Append(new Controls(controlsXml));
                worksheetPart.Worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document["Controls"].CellValue(3, 1, "After");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
                OleObjects oleObjects = Assert.Single(worksheet.Elements<OleObjects>());
                Controls controls = Assert.Single(worksheet.Elements<Controls>());

                Assert.Contains("rIdOlePackage", oleObjects.OuterXml);
                Assert.Contains("ApproveButton", controls.OuterXml);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "OLE objects"
                    && feature.Count == 1
                    && feature.Details.Any(detail => detail.Contains("Controls", StringComparison.OrdinalIgnoreCase)));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Form controls"
                    && feature.Count == 1
                    && feature.Details.Any(detail => detail.Contains("Controls", StringComparison.OrdinalIgnoreCase)));
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_FeatureReport_FailsFastForBlockedFeatures() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.FeatureReportFailFast.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Links");
                sheet.CellValue(1, 1, "Resource");
                sheet.SetHyperlink(2, 1, "https://example.org/spec", display: "Spec");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();

                Assert.Empty(report.FindFeatures("VBA macros"));
                ExcelFeatureFinding externalLinks = Assert.Single(report.FindFeatures("External hyperlinks"));
                Assert.Equal(ExcelFeatureSupportLevel.PartiallyEditable, externalLinks.SupportLevel);
                Assert.Same(report, report.EnsureNoFeatures("VBA macros"));

                InvalidOperationException namedException = Assert.Throws<InvalidOperationException>(
                    () => report.EnsureNoFeatures("External hyperlinks", "VBA macros"));
                Assert.Contains("External hyperlinks", namedException.Message);
                Assert.Contains("https://example.org/spec", namedException.Message);

                InvalidOperationException levelException = Assert.Throws<InvalidOperationException>(
                    () => report.EnsureNoFeatures(ExcelFeatureSupportLevel.PartiallyEditable));
                Assert.Contains("PartiallyEditable", levelException.Message);
                Assert.Contains("External hyperlinks", levelException.Message);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_RoundTrip_PreservesExternalLinksAndCustomXml() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.PreserveExternalLinksCustomXml.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Links");
                sheet.CellValue(1, 1, "Resource");
                sheet.CellValue(2, 1, "Spec");
                sheet.SetHyperlink(2, 1, "https://example.org/spec", display: "Spec");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                CustomXmlPart customXmlPart = spreadsheet.WorkbookPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using var stream = new MemoryStream(Encoding.UTF8.GetBytes("<metadata><owner>OfficeIMO</owner></metadata>"));
                customXmlPart.FeedData(stream);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document["Links"].CellValue(3, 1, "Edited");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                Assert.Single(spreadsheet.WorkbookPart!.CustomXmlParts);
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.Single();
                HyperlinkRelationship relationship = Assert.Single(worksheetPart.HyperlinkRelationships);
                Assert.Equal(new Uri("https://example.org/spec"), relationship.Uri);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "External hyperlinks"
                    && feature.Details.Any(detail => detail.Contains("https://example.org/spec", StringComparison.OrdinalIgnoreCase)));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Custom XML parts"
                    && feature.Details.Any(detail => detail.Contains("/customXml/", StringComparison.OrdinalIgnoreCase)));
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_RoundTrip_PreservesMacroAndEmbeddedPackageParts() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.PreserveMacroEmbeddedPackage.xlsx");
            byte[] vbaBytes = Encoding.ASCII.GetBytes("OfficeIMO macro project placeholder");
            byte[] embeddedBytes = Encoding.ASCII.GetBytes("OfficeIMO embedded workbook placeholder");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Package");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Before");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                VbaProjectPart vbaProjectPart = workbookPart.AddNewPart<VbaProjectPart>();
                using (var stream = new MemoryStream(vbaBytes)) {
                    vbaProjectPart.FeedData(stream);
                }

                WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
                EmbeddedPackagePart embeddedPackagePart = worksheetPart.AddEmbeddedPackagePart(EmbeddedPackagePartType.Xlsx);
                using var embeddedStream = new MemoryStream(embeddedBytes);
                embeddedPackagePart.FeedData(embeddedStream);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document["Package"].CellValue(3, 1, "After");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                Assert.NotNull(workbookPart.VbaProjectPart);
                using (var stream = workbookPart.VbaProjectPart!.GetStream(FileMode.Open, FileAccess.Read)) {
                    using var buffer = new MemoryStream();
                    stream.CopyTo(buffer);
                    Assert.Equal(vbaBytes, buffer.ToArray());
                }

                WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
                EmbeddedPackagePart embeddedPackagePart = Assert.Single(worksheetPart.EmbeddedPackageParts);
                using (var stream = embeddedPackagePart.GetStream(FileMode.Open, FileAccess.Read)) {
                    using var buffer = new MemoryStream();
                    stream.CopyTo(buffer);
                    Assert.Equal(embeddedBytes, buffer.ToArray());
                }
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "VBA macros"
                    && feature.Details.Any(detail => detail.Contains("vbaProject", StringComparison.OrdinalIgnoreCase)));
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Embedded packages"
                    && feature.Details.Any(detail => detail.Contains("/embeddings/", StringComparison.OrdinalIgnoreCase)));
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_RoundTrip_PreservesDigitalSignatureMetadataParts() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.PreserveDigitalSignatureMetadata.xlsx");
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo /></Signature>");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Signed");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(2, 1, "Before");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                spreadsheet.AddDigitalSignatureOriginPart();
                DigitalSignatureOriginPart originPart = spreadsheet.DigitalSignatureOriginPart!;
                XmlSignaturePart signaturePart = originPart.AddNewPart<XmlSignaturePart>();
                using (var stream = new MemoryStream(signatureBytes)) {
                    signaturePart.FeedData(stream);
                }

                ExtendedFilePropertiesPart appPart = spreadsheet.ExtendedFilePropertiesPart ?? spreadsheet.AddExtendedFilePropertiesPart();
                appPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                appPart.Properties.DigitalSignature = new DocumentFormat.OpenXml.ExtendedProperties.DigitalSignature();
                appPart.Properties.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document["Signed"].CellValue(3, 1, "After");
                ExcelSignatureInfo signatures = document.InspectSignatures();
                Assert.True(signatures.HasSignatures);
                Assert.True(signatures.HasDigitalSignatureOriginPart);
                Assert.Equal(1, signatures.XmlSignaturePartCount);
                Assert.True(signatures.HasApplicationSignatureMetadata);
                ExcelSignedWorkbookMutationException blocked = Assert.Throws<ExcelSignedWorkbookMutationException>(() =>
                    document.Save());
                Assert.True(blocked.SignatureInfo.HasSignatures);
                document.Save(new ExcelSaveOptions {
                    SignatureMutationPolicy = ExcelSignatureMutationPolicy.PreserveSignatureMarkup
                });
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                Assert.NotNull(spreadsheet.DigitalSignatureOriginPart);
                DigitalSignatureOriginPart originPart = spreadsheet.DigitalSignatureOriginPart!;
                XmlSignaturePart signaturePart = Assert.Single(originPart.XmlSignatureParts);
                using (Stream stream = signaturePart.GetStream(FileMode.Open, FileAccess.Read)) {
                    using var buffer = new MemoryStream();
                    stream.CopyTo(buffer);
                    Assert.Equal(signatureBytes, buffer.ToArray());
                }

                Assert.NotNull(spreadsheet.ExtendedFilePropertiesPart?.Properties?.DigitalSignature);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelFeatureReport report = document.InspectFeatures();
                ExcelFeatureFinding signatures = Assert.Single(report.FindFeatures("Digital signatures"));
                Assert.Equal(ExcelFeatureSupportLevel.Preserved, signatures.SupportLevel);
                Assert.Contains(signatures.Details, detail => detail.Contains("/origin.sigs", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(signatures.Details, detail => detail.Contains("/_xmlsignatures/", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(signatures.Details, detail => detail.Contains("extended application properties", StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculationPolicy_IgnoresUnsupportedOrOversizedFormulas() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculationUnsupported.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Calc");
                sheet.CellValue(1, 1, 2d);
                sheet.CellFormula(2, 1, "VLOOKUP(A1,B1:C2,2,FALSE)");
                sheet.CellFormula(3, 1, "SUM(" + new string('A', 9000) + ")");
                sheet.CellFormula(4, 1, "SUM(Calc!A1:A2)");

                Assert.Equal(0, document.RecalculateSupportedFormulas());
                Assert.False(sheet.TryGetCachedFormulaValue(2, 1, out _));
                Assert.False(sheet.TryGetCachedFormulaValue(3, 1, out _));
                Assert.False(sheet.TryGetCachedFormulaValue(4, 1, out _));

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_Sort_UsesNumericFormulaCachesAndKeepsMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.SortFormulaCaches.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Sort");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Score");
                sheet.CellValue(2, 1, "Hundred");
                sheet.CellValue(2, 3, 100d);
                sheet.CellFormula(2, 2, "C2+0");
                sheet.CellValue(3, 1, "Two");
                sheet.CellValue(3, 3, 2d);
                sheet.CellFormula(3, 2, "C3+0");
                sheet.SetComment(3, 1, "moves with Two", author: "Tester");

                Assert.Equal(2, document.RecalculateSupportedFormulas());
                sheet.SortRangeByColumn("A1:B3", 2, ascending: true, hasHeader: true);

                Assert.Equal("Two", sheet.CellAt(2, 1).GetValue<string>());
                Assert.True(sheet.HasComment(2, 1));
                Assert.Equal("Hundred", sheet.CellAt(3, 1).GetValue<string>());

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_Sort_SplitsNonContiguousHyperlinkRemaps() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.SortHyperlinkRemap.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Sort");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Score");
                sheet.CellValue(2, 1, "Hundred");
                sheet.CellValue(2, 2, 100d);
                sheet.CellValue(3, 1, "Two");
                sheet.CellValue(3, 2, 2d);
                sheet.CellValue(4, 1, "Fifty");
                sheet.CellValue(4, 2, 50d);
                sheet.SetHyperlink(2, 1, "https://example.com", display: "Linked", style: false);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Hyperlink hyperlink = worksheetPart.Worksheet.Descendants<Hyperlink>().Single();
                hyperlink.Reference = "A2:A3";
                worksheetPart.Worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets[0];
                sheet.SortRangeByColumn("A1:B4", 2, ascending: true, hasHeader: true);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                string[] references = worksheetPart.Worksheet.Descendants<Hyperlink>()
                    .Select(hyperlink => hyperlink.Reference?.Value ?? string.Empty)
                    .OrderBy(reference => reference)
                    .ToArray();

                Assert.Equal(new[] { "A2", "A4" }, references);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_WorkbookAndWorksheetProtection_PreserveLegacyHashes() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.Protection.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Protected");
                sheet.CellValue(1, 1, "locked");

                document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                    ProtectStructure = true,
                    ProtectWindows = true,
                    LegacyPasswordHash = "CAFE"
                });
                sheet.Protect(new ExcelSheetProtectionOptions {
                    AllowSort = true,
                    LegacyPasswordHash = "BEEF"
                });
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Create(Path.Combine(_directoryWithFiles, "ClosedXmlGap.InvalidProtectionHash.xlsx"))) {
                Assert.Throws<ArgumentException>(() => document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                    LegacyPasswordHash = "NOPE"
                }));
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorkbookProtection workbookProtection = spreadsheet.WorkbookPart!.Workbook.GetFirstChild<WorkbookProtection>()!;
                Assert.True(workbookProtection.LockStructure!.Value);
                Assert.True(workbookProtection.LockWindows!.Value);
                Assert.Equal("CAFE", workbookProtection.WorkbookPassword!.Value);

                WorksheetPart worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                SheetProtection sheetProtection = worksheetPart.Worksheet.Elements<SheetProtection>().First();
                Assert.True(sheetProtection.Sheet!.Value);
                Assert.False(sheetProtection.Sort!.Value);
                Assert.Equal("BEEF", sheetProtection.Password!.Value);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.True(document.IsWorkbookProtected);
                Assert.True(document.Sheets[0].IsProtected);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculationProperties_AreInsertedBeforeLaterWorkbookNodes() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculationOrder.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Calc");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Workbook workbook = spreadsheet.WorkbookPart!.Workbook;
                workbook.Append(new PivotCaches());
                workbook.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document.ConfigureFullCalculationOnOpen();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                Workbook workbook = spreadsheet.WorkbookPart!.Workbook;
                var children = workbook.ChildElements.ToList();
                int calculationIndex = children.FindIndex(element => element is CalculationProperties);
                int pivotCachesIndex = children.FindIndex(element => element is PivotCaches);

                Assert.True(calculationIndex >= 0);
                Assert.True(pivotCachesIndex >= 0);
                Assert.True(calculationIndex < pivotCachesIndex);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_FullCalculationOnOpen_PreservesManualCalculationMode() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculationMode.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Manual");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Workbook workbook = spreadsheet.WorkbookPart!.Workbook;
                workbook.Append(new CalculationProperties { CalculationMode = CalculateModeValues.Manual });
                workbook.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document.ConfigureFullCalculationOnOpen();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                CalculationProperties calculationProperties = spreadsheet.WorkbookPart!.Workbook.GetFirstChild<CalculationProperties>()!;

                Assert.Equal(CalculateModeValues.Manual, calculationProperties.CalculationMode!.Value);
                Assert.True(calculationProperties.ForceFullCalculation!.Value);
                Assert.True(calculationProperties.FullCalculationOnLoad!.Value);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_Range_AcceptsSingleCellAddress() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.SingleCellRange.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Range");
                ExcelRange range = sheet.Range("A1");
                ExcelRange absoluteCellRange = sheet.Range("$B$2");
                ExcelRange absoluteMultiCellRange = sheet.Range("$C$3:$D$4");

                Assert.Equal("A1:A1", range.Address);
                Assert.Equal("B2:B2", absoluteCellRange.Address);
                Assert.Equal("C3:D4", absoluteMultiCellRange.Address);
                range.FirstCell.SetValue("single");
                Assert.Equal("single", range.FirstCell.GetValue<string>());

                range.Clear(ExcelClearOptions.Values);
                Assert.True(range.FirstCell.GetValue().IsBlank);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_RangeCreateTable_ReturnsResolvedTableName() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ResolvedTableName.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Table");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Score");
                sheet.CellValue(2, 1, "Alpha");
                sheet.CellValue(2, 2, 10d);

                ExcelTable table = sheet.Range("A1:B2").CreateTable("My Table");
                Assert.Equal("My_Table", table.NameOrRange);
                Assert.Equal("A1:B2", table.Range);
                table.SetStyle(ExcelTableStyle.TableStyleMedium4);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_Inspection_HonorsSheetProtectionDefaults() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ProtectionDefaults.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Defaults");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                worksheetPart.Worksheet.Append(new SheetProtection { Sheet = true });
                worksheetPart.Worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                ExcelWorksheetSnapshot worksheet = Assert.Single(snapshot.Worksheets);

                Assert.NotNull(worksheet.Protection);
                Assert.True(worksheet.Protection!.AllowSelectLockedCells);
                Assert.True(worksheet.Protection.AllowSelectUnlockedCells);
                Assert.False(worksheet.Protection.AllowFormatCells);
                Assert.False(worksheet.Protection.AllowSort);
                Assert.False(worksheet.Protection.AllowAutoFilter);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_GetRichText_DoesNotMaterializeMissingCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.RichTextReadMissingCell.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("RichText");

                Assert.Empty(sheet.CellAt(10, 3).GetRichText());

                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                string[] references = worksheetPart.Worksheet.Descendants<Cell>()
                    .Select(cell => cell.CellReference?.Value ?? string.Empty)
                    .ToArray();

                Assert.DoesNotContain("C10", references);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ClearHyperlinks_PreservesNonOverlappingRangeSegments() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ClearHyperlinkSegments.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Links");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "One");
                sheet.CellValue(3, 1, "Two");
                sheet.CellValue(4, 1, "Three");
                sheet.SetHyperlink(2, 1, "https://example.com", display: "Linked", style: false);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Hyperlink hyperlink = worksheetPart.Worksheet.Descendants<Hyperlink>().Single();
                hyperlink.Reference = "A2:A4";
                worksheetPart.Worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document.Sheets[0].Range("A3").Clear(ExcelClearOptions.Hyperlinks);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                string[] references = worksheetPart.Worksheet.Descendants<Hyperlink>()
                    .Select(hyperlink => hyperlink.Reference?.Value ?? string.Empty)
                    .OrderBy(reference => reference)
                    .ToArray();

                Assert.Equal(new[] { "A2", "A4" }, references);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ClearHyperlinks_NoOverlapDoesNotSplitReferenceList() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ClearHyperlinkNoOverlap.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Links");
                sheet.CellValue(2, 1, "One");
                sheet.CellValue(4, 1, "Two");
                sheet.SetHyperlink(2, 1, "https://example.com", display: "Linked", style: false);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Hyperlink hyperlink = worksheetPart.Worksheet.Descendants<Hyperlink>().Single();
                hyperlink.Reference = "A2 A4";
                worksheetPart.Worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document.Sheets[0].Range("C1:C3").Clear(ExcelClearOptions.Hyperlinks);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Hyperlink hyperlink = worksheetPart.Worksheet.Descendants<Hyperlink>().Single();
                Assert.Equal("A2 A4", hyperlink.Reference?.Value);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_Sort_RewritesRelativeFormulaRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.SortFormulaRows.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Sort");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(1, 2, "Formula");
                sheet.CellValue(2, 1, 2d);
                sheet.CellValue(3, 1, 1d);
                sheet.CellAt(2, 2).SetFormula("A2");
                sheet.CellAt(3, 2).SetFormula("A3");

                sheet.SortRangeByColumn("A1:B3", 1, ascending: true, hasHeader: true);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Dictionary<string, string?> formulas = worksheetPart.Worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellFormula != null)
                    .ToDictionary(cell => cell.CellReference!.Value!, cell => cell.CellFormula?.Text);

                Assert.Equal("A2", formulas["B2"]);
                Assert.Equal("A3", formulas["B3"]);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_Sort_DoesNotMaterializeSparseBlankCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.SortSparse.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Sparse");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(1, 2, "Blank");
                sheet.CellValue(2, 1, 2d);
                sheet.CellValue(3, 1, 1d);

                sheet.SortRangeByColumn("A1:B3", 1, ascending: true, hasHeader: true);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                string[] references = worksheetPart.Worksheet.Descendants<Cell>()
                    .Select(cell => cell.CellReference?.Value ?? string.Empty)
                    .ToArray();

                Assert.DoesNotContain("B2", references);
                Assert.DoesNotContain("B3", references);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_WorkbookProtection_IsInsertedBeforeBookViews() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ProtectionBookViews.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Protected");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Workbook workbook = spreadsheet.WorkbookPart!.Workbook;
                if (workbook.GetFirstChild<BookViews>() == null) {
                    workbook.InsertBefore(new BookViews(new WorkbookView()), workbook.GetFirstChild<Sheets>());
                    workbook.Save();
                }
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document.ProtectWorkbook(new ExcelWorkbookProtectionOptions { LegacyPasswordHash = "CAFE" });
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                Workbook workbook = spreadsheet.WorkbookPart!.Workbook;
                var children = workbook.ChildElements.ToList();
                int protectionIndex = children.FindIndex(element => element is WorkbookProtection);
                int bookViewsIndex = children.FindIndex(element => element is BookViews);

                Assert.True(protectionIndex >= 0);
                Assert.True(bookViewsIndex >= 0);
                Assert.True(protectionIndex < bookViewsIndex);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ClearRange_None_DoesNotMaterializeCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ClearNone.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Clear");
                sheet.ClearRange("C5:D6", ExcelClearOptions.None);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Assert.Empty(worksheetPart.Worksheet.Descendants<Cell>());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ClearRange_Values_DoesNotMaterializeSparseBlankCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ClearSparseValues.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Clear");
                sheet.CellValue(1, 1, "Keep");
                sheet.CellValue(5, 5, "Clear");
                sheet.ClearRange("A1:E5", ExcelClearOptions.Values);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                string[] references = worksheetPart.Worksheet.Descendants<Cell>()
                    .Select(cell => cell.CellReference?.Value ?? string.Empty)
                    .OrderBy(reference => reference)
                    .ToArray();

                Assert.Equal(new[] { "A1", "E5" }, references);
                Assert.All(worksheetPart.Worksheet.Descendants<Cell>(), cell => Assert.Null(cell.CellValue));
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ClearRange_CellFields_ScansExistingSparseCellsOnly() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ClearSparseCellFields.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Clear");
                sheet.CellValue(1, 1, "Keep");
                sheet.CellAt(10, 5).SetValue("Clear").SetFillColor("FFF2CC");
                sheet.CellAt(10, 6).SetFormula("E10");
                sheet.CellAt(12, 5).SetValue("Outside").SetFillColor("D9EAD3");

                sheet.ClearRange("E10:F10", ExcelClearOptions.Values | ExcelClearOptions.Formulas | ExcelClearOptions.Styles);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value != null)
                    .ToDictionary(cell => cell.CellReference!.Value!);

                Assert.Equal(new[] { "A1", "E10", "E12", "F10" }, cells.Keys.OrderBy(reference => reference).ToArray());
                Assert.Null(cells["E10"].CellValue);
                Assert.Null(cells["E10"].StyleIndex);
                Assert.Null(cells["F10"].CellFormula);
                Assert.Null(cells["F10"].CellValue);
                Assert.NotNull(cells["E12"].CellValue);
                Assert.NotNull(cells["E12"].StyleIndex);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ClearRange_Comments_RemovesOnlyShapesInsideSparseRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ClearSparseComments.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Clear");
                sheet.SetComment(1, 1, "Remove first", author: "Tester");
                sheet.SetComment(5, 5, "Remove second", author: "Tester");
                sheet.SetComment(7, 7, "Keep outside", author: "Tester");

                sheet.ClearRange("A1:E5", ExcelClearOptions.Comments);

                Assert.False(sheet.HasComment(1, 1));
                Assert.False(sheet.HasComment(5, 5));
                Assert.True(sheet.HasComment(7, 7));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                string[] commentReferences = worksheetPart.WorksheetCommentsPart!.Comments!.CommentList!.Elements<Comment>()
                    .Select(comment => comment.Reference?.Value ?? string.Empty)
                    .ToArray();

                Assert.Equal(new[] { "G7" }, commentReferences);

                VmlDrawingPart vmlPart = Assert.Single(worksheetPart.VmlDrawingParts);
                XDocument vml = XDocument.Load(vmlPart.GetStream());
                XNamespace excelNamespace = "urn:schemas-microsoft-com:office:excel";
                string[] vmlCoordinates = vml.Root!.Descendants(excelNamespace + "ClientData")
                    .Select(clientData => string.Join(
                        ",",
                        clientData.Element(excelNamespace + "Row")?.Value.Trim(),
                        clientData.Element(excelNamespace + "Column")?.Value.Trim()))
                    .ToArray();

                Assert.Equal(new[] { "6,6" }, vmlCoordinates);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ClearRange_Comments_NoOverlapPreservesArtifacts() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ClearCommentsNoOverlap.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Clear");
                sheet.SetComment(1, 1, "Keep", author: "Tester");

                sheet.ClearRange("C3:D4", ExcelClearOptions.Comments);
                Assert.True(sheet.HasComment(1, 1));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                string commentReference = Assert.Single(worksheetPart.WorksheetCommentsPart!.Comments!.CommentList!.Elements<Comment>())
                    .Reference?.Value ?? string.Empty;

                Assert.Equal("A1", commentReference);

                VmlDrawingPart vmlPart = Assert.Single(worksheetPart.VmlDrawingParts);
                XDocument vml = XDocument.Load(vmlPart.GetStream());
                XNamespace excelNamespace = "urn:schemas-microsoft-com:office:excel";
                string coordinate = Assert.Single(vml.Root!.Descendants(excelNamespace + "ClientData")
                    .Select(clientData => string.Join(
                        ",",
                        clientData.Element(excelNamespace + "Row")?.Value.Trim(),
                        clientData.Element(excelNamespace + "Column")?.Value.Trim())));

                Assert.Equal("0,0", coordinate);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ClearComment_MissingCellPreservesOtherCommentArtifacts() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ClearMissingComment.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Clear");
                sheet.SetComment(1, 1, "Keep", author: "Tester");

                sheet.ClearComment("C3");
                Assert.True(sheet.HasComment(1, 1));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                string commentReference = Assert.Single(worksheetPart.WorksheetCommentsPart!.Comments!.CommentList!.Elements<Comment>())
                    .Reference?.Value ?? string.Empty;

                Assert.Equal("A1", commentReference);

                VmlDrawingPart vmlPart = Assert.Single(worksheetPart.VmlDrawingParts);
                XDocument vml = XDocument.Load(vmlPart.GetStream());
                XNamespace excelNamespace = "urn:schemas-microsoft-com:office:excel";
                string coordinate = Assert.Single(vml.Root!.Descendants(excelNamespace + "ClientData")
                    .Select(clientData => string.Join(
                        ",",
                        clientData.Element(excelNamespace + "Row")?.Value.Trim(),
                        clientData.Element(excelNamespace + "Column")?.Value.Trim())));

                Assert.Equal("0,0", coordinate);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ClearArrayFormula_ClearsEntireSpillRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ClearArraySpill.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Array");
                sheet.CellValue(1, 1, 10d);
                sheet.CellValue(1, 2, 20d);
                sheet.CellValue(2, 1, 30d);
                sheet.CellValue(2, 2, 40d);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Cell anchor = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
                anchor.CellFormula = new CellFormula("SUM(C1:C2)") {
                    FormulaType = CellFormulaValues.Array,
                    Reference = "A1:B2"
                };
                worksheetPart.Worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document.Sheets[0].ClearArrayFormula("B2");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellReference?.Value is "A1" or "B1" or "A2" or "B2")
                    .ToDictionary(cell => cell.CellReference!.Value!);

                Assert.All(cells.Values, cell => {
                    Assert.Null(cell.CellFormula);
                    Assert.Null(cell.CellValue);
                });
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_FormulaReadApis_DoNotMaterializeMissingCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.FormulaReadMissingCell.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Formula");

                Assert.Null(sheet.GetFormulaText(10, 3));
                Assert.False(sheet.TryGetCachedFormulaValue(10, 3, out string? cachedValue));
                Assert.Null(cachedValue);

                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                string[] references = worksheetPart.Worksheet.Descendants<Cell>()
                    .Select(cell => cell.CellReference?.Value ?? string.Empty)
                    .ToArray();

                Assert.DoesNotContain("C10", references);
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ConditionalFormattingAndDataValidation_Management() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.Rules.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Rules");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, 10d);
                sheet.CellValue(3, 1, 20d);
                sheet.CellValue(4, 1, 20d);

                sheet.AddConditionalFormulaRule("A2:A4", "A2>15", stopIfTrue: true);
                sheet.AddConditionalDuplicateValuesRule("A2:A4");
                sheet.AddConditionalTopBottomRule("A2:A4", 1);

                IReadOnlyList<ExcelConditionalFormattingInfo> rules = sheet.GetConditionalFormattingRules("A2:A4");
                Assert.Equal(3, rules.Count);
                Assert.Contains(rules, rule => rule.Type == "Expression" && rule.StopIfTrue);
                Assert.Contains(rules, rule => rule.Type == "DuplicateValues");
                Assert.Contains(rules, rule => rule.Type == "Top10");

                sheet.ValidationWholeNumber("B2:B4", DataValidationOperatorValues.Between, 1, 30);
                sheet.SetDataValidationMessages("B2:B4", new ExcelDataValidationMessageOptions {
                    PromptTitle = "Allowed",
                    Prompt = "1-30 only",
                    ErrorTitle = "Invalid",
                    Error = "Choose a number between 1 and 30"
                });

                ExcelDataValidationInfo validation = Assert.Single(sheet.GetDataValidations("B3:B3"));
                Assert.Equal("Allowed", validation.PromptTitle);
                Assert.Equal("Invalid", validation.ErrorTitle);

                sheet.RemoveDataValidations("B3:B3");
                Assert.Single(sheet.GetDataValidations("B2:B2"));
                Assert.Empty(sheet.GetDataValidations("B3:B3"));
                Assert.Single(sheet.GetDataValidations("B4:B4"));

                sheet.RemoveDataValidations("B2:B4");
                Assert.Empty(sheet.GetDataValidations("B2:B4"));

                sheet.ClearConditionalFormatting("A3:A3");
                Assert.Equal(3, sheet.GetConditionalFormattingRules("A2:A2").Count);
                Assert.Empty(sheet.GetConditionalFormattingRules("A3:A3"));
                Assert.Equal(3, sheet.GetConditionalFormattingRules("A4:A4").Count);

                sheet.ClearConditionalFormatting("A2:A4");
                Assert.Empty(sheet.GetConditionalFormattingRules("A2:A4"));

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_RangeMetadata_NoOverlapClearPreservesReferenceLists() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.RulesNoOverlap.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Rules");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, 10d);
                sheet.CellValue(3, 1, 20d);
                sheet.CellValue(4, 1, 30d);
                sheet.CellValue(6, 1, 40d);
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 2, 20d);
                sheet.CellValue(4, 2, 30d);
                sheet.CellValue(6, 2, 40d);

                sheet.AddConditionalFormulaRule("A2:A4 A6", "A2>15");
                sheet.ValidationWholeNumber("B2:B4 B6", DataValidationOperatorValues.Between, 1, 50);

                Assert.Single(sheet.GetConditionalFormattingRules("A6"));
                Assert.Single(sheet.GetDataValidations("B6"));

                sheet.SetDataValidationMessages("D1:D2", new ExcelDataValidationMessageOptions {
                    PromptTitle = "Outside",
                    Prompt = "Should not apply",
                    ErrorTitle = "Outside",
                    Error = "Should not apply"
                });
                ExcelDataValidationInfo untouchedValidation = Assert.Single(sheet.GetDataValidations("B2:B4"));
                Assert.Null(untouchedValidation.PromptTitle);
                Assert.Null(untouchedValidation.ErrorTitle);

                sheet.ClearRange("D1:D3", ExcelClearOptions.ConditionalFormatting | ExcelClearOptions.DataValidations);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Assert.Equal("A2:A4 A6", worksheetPart.Worksheet.Elements<ConditionalFormatting>().Single().SequenceOfReferences?.InnerText);
                Assert.Equal("B2:B4 B6", worksheetPart.Worksheet.Descendants<DataValidation>().Single().SequenceOfReferences?.InnerText);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}

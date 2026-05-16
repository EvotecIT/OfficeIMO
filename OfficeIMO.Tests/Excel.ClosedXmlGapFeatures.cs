using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using ExcelTableStyle = OfficeIMO.Excel.TableStyle;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ClosedXmlGap_ObjectModel_RichText_Sort_Table_And_Clear() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ObjectModel.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculationPolicy_EvaluatesSupportedFormulasBeforeSave() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.Calculation.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Calc");
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
        public void Test_ClosedXmlGap_CalculationPolicy_IgnoresUnsupportedOrOversizedFormulas() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculationUnsupported.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Calc");
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_Sort_UsesNumericFormulaCachesAndKeepsMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.SortFormulaCaches.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sort");
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_Sort_SplitsNonContiguousHyperlinkRemaps() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.SortHyperlinkRemap.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sort");
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_WorkbookAndWorksheetProtection_PreserveLegacyHashes() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.Protection.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Protected");
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.True(document.IsWorkbookProtected);
                Assert.True(document.Sheets[0].IsProtected);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_CalculationProperties_AreInsertedBeforeLaterWorkbookNodes() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.CalculationOrder.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Calc");
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
        public void Test_ClosedXmlGap_Range_AcceptsSingleCellAddress() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.SingleCellRange.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Range");
                ExcelRange range = sheet.Range("A1");

                Assert.Equal("A1:A1", range.Address);
                range.FirstCell.SetValue("single");
                Assert.Equal("single", range.FirstCell.GetValue<string>());

                range.Clear(ExcelClearOptions.Values);
                Assert.True(range.FirstCell.GetValue().IsBlank);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_RangeCreateTable_ReturnsResolvedTableName() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ResolvedTableName.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Table");
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ClearHyperlinks_PreservesNonOverlappingRangeSegments() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ClearHyperlinkSegments.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Links");
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_Sort_RewritesRelativeFormulaRows() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.SortFormulaRows.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sort");
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_Sort_DoesNotMaterializeSparseBlankCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.SortSparse.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Sparse");
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_WorkbookProtection_IsInsertedBeforeBookViews() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ProtectionBookViews.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Protected");
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ClearRange_None_DoesNotMaterializeCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.ClearNone.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Clear");
                sheet.ClearRange("C5:D6", ExcelClearOptions.None);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Assert.Empty(worksheetPart.Worksheet.Descendants<Cell>());
            }
        }

        [Fact]
        public void Test_ClosedXmlGap_ConditionalFormattingAndDataValidation_Management() {
            string filePath = Path.Combine(_directoryWithFiles, "ClosedXmlGap.Rules.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Rules");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, 10d);
                sheet.CellValue(3, 1, 20d);
                sheet.CellValue(4, 1, 20d);

                sheet.AddConditionalFormulaRule("A2:A4", "A2>15", stopIfTrue: true);
                sheet.AddConditionalDuplicateValuesRule("A2:A4");
                sheet.AddConditionalTopBottomRule("A2:A4", 1);

                IReadOnlyList<ExcelConditionalFormattingInfo> rules = sheet.GetConditionalFormattingRules("A2:A4");
                Assert.Equal(3, rules.Count);
                Assert.Contains(rules, rule => rule.Type == ConditionalFormatValues.Expression.ToString() && rule.StopIfTrue);
                Assert.Contains(rules, rule => rule.Type == ConditionalFormatValues.DuplicateValues.ToString());
                Assert.Contains(rules, rule => rule.Type == ConditionalFormatValues.Top10.ToString());

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

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}

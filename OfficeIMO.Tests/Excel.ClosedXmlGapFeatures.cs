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

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorkbookProtection workbookProtection = spreadsheet.WorkbookPart!.Workbook.GetFirstChild<WorkbookProtection>()!;
                Assert.True(workbookProtection.LockStructure!.Value);
                Assert.True(workbookProtection.LockWindows!.Value);
                Assert.Equal("CAFE", workbookProtection.WorkbookPassword!.Value);

                WorksheetPart worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                SheetProtection sheetProtection = worksheetPart.Worksheet.Elements<SheetProtection>().First();
                Assert.True(sheetProtection.Sheet!.Value);
                Assert.True(sheetProtection.Sort!.Value);
                Assert.Equal("BEEF", sheetProtection.Password!.Value);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.True(document.IsWorkbookProtected);
                Assert.True(document.Sheets[0].IsProtected);
                Assert.Empty(document.ValidateOpenXml());
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

                sheet.RemoveDataValidations("B2:B4");
                Assert.Empty(sheet.GetDataValidations("B2:B4"));
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

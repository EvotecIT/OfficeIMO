using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_WorksheetValuesCopy_StreamsUnopenedSourceWhileCopyingRuleNames(bool merge) {
            string sourcePath = Path.Combine(_directoryWithFiles, "ValuesCopyStreamedSource-" + merge + ".xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ValuesCopyStreamedTarget-" + merge + ".xlsx");
            CreateLayoutSource(sourcePath);
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, true)) {
                SheetData sheetData = GetWorksheetPartByNameForOperations(source, "Data").Worksheet!.GetFirstChild<SheetData>()!;
                for (uint row = 5U; row <= 3000U; row++) {
                    sheetData.Append(new Row(new Cell { CellReference = "A" + row, CellValue = new CellValue((double)row) }) { RowIndex = row });
                }
            }

            using (var createdTarget = ExcelDocument.Create(targetPath)) {
                createdTarget.AddSheet("Existing").Cell(1, 1).SetValue("Existing");
                createdTarget.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                WorksheetPart sourcePart = sourceDocument.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single();
                if (merge) {
                    targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions {
                        TargetSheetNames = new Dictionary<string, string> { ["Data"] = "Copy" }
                    });
                } else {
                    targetDocument.CopyWorksheetFrom(sourceDocument, "Data", "Copy");
                }

                Assert.False(sourcePart.IsRootElementLoaded);
                targetDocument.Save();
            }

            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                Worksheet copy = GetWorksheetPartByNameForOperations(target, "Copy").Worksheet!;
                Assert.Equal("AllowedValues", copy.Descendants<DataValidation>().Single().GetFirstChild<Formula1>()!.Text);
                DefinedName allowed = target.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "AllowedValues");
                Assert.Equal("Copy!$H$1:$H$2", allowed.Text.Replace("'", string.Empty));
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            using (var reloaded = ExcelDocument.Load(targetPath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.True(reloaded["Copy"].TryGetCellText(3000, 1, out string last));
                Assert.Equal("3000", last);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void Test_WorksheetValuesCopy_ExternalWorkbookRulesFollowPackageModePolicy(bool merge, bool copyExternalWorkbookReferences) {
            string sourcePath = Path.Combine(_directoryWithFiles, $"ValuesCopyExternalSource-{merge}-{copyExternalWorkbookReferences}.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, $"ValuesCopyExternalTarget-{merge}-{copyExternalWorkbookReferences}.xlsx");
            using (var createdSource = ExcelDocument.Create(sourcePath)) {
                ExcelSheet data = createdSource.AddSheet("Data");
                data.Cell(1, 1).SetValue(1);
                data.Cell(2, 1).SetValue(2);
                createdSource.Save();
            }

            AddLinkedWorkbook(sourcePath, "Prices.xlsx");
            AddDefinedName(sourcePath, "ExternalValue", "[1]Sheet1!$A$1");
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, true)) {
                Worksheet worksheet = GetWorksheetPartByNameForOperations(source, "Data").Worksheet!;
                worksheet.InsertAfter(new ConditionalFormatting(
                    new ConditionalFormattingRule(new Formula("A1>ExternalValue")) { Type = ConditionalFormatValues.Expression, Priority = 1 },
                    new ConditionalFormattingRule(new Formula("A1>[1]Sheet1!$A$2")) { Type = ConditionalFormatValues.Expression, Priority = 2 }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A2" }
                }, worksheet.GetFirstChild<SheetData>()!);
                worksheet.Save();
            }

            using (var createdTarget = ExcelDocument.Create(targetPath)) {
                createdTarget.AddSheet("Existing").Cell(1, 1).SetValue("Existing");
                createdTarget.Save();
            }

            // The target already links one workbook, so the copied link becomes [2].
            AddLinkedWorkbook(targetPath, "Rates.xlsx");

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                Action copy = merge
                    ? () => targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions {
                        TargetSheetNames = new Dictionary<string, string> { ["Data"] = "Copy" },
                        CopyExternalWorkbookReferences = copyExternalWorkbookReferences
                    })
                    : () => targetDocument.CopyWorksheetFrom(sourceDocument, "Data", "Copy", ExcelSheetNameValidationMode.Strict,
                        new ExcelWorksheetCopyOptions { CopyExternalWorkbookReferences = copyExternalWorkbookReferences });
                if (!copyExternalWorkbookReferences) {
                    Assert.Throws<InvalidOperationException>(copy);
                    Assert.Equal(new[] { "Existing" }, targetDocument.Sheets.Select(sheet => sheet.Name));
                    File.Delete(sourcePath);
                    return;
                }

                copy();
                targetDocument.Save();
            }

            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                Workbook workbook = target.WorkbookPart!.Workbook!;
                Assert.Equal(2, workbook.ExternalReferences!.Elements<ExternalReference>().Count());
                DefinedName externalValue = workbook.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "ExternalValue");
                Assert.Equal("[2]Sheet1!$A$1", externalValue.Text);
                Assert.Equal(new[] { "A1>ExternalValue", "A1>[2]Sheet1!$A$2" },
                    GetWorksheetPartByNameForOperations(target, "Copy").Worksheet!.Descendants<ConditionalFormattingRule>()
                        .Select(rule => rule.GetFirstChild<Formula>()!.Text));
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            using (var reloaded = ExcelDocument.Load(targetPath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.True(reloaded["Copy"].TryGetCellText(2, 1, out string value));
                Assert.Equal("2", value);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Theory]
        [InlineData(false, ExcelWorksheetCopyMode.Values)]
        [InlineData(false, ExcelWorksheetCopyMode.Package)]
        [InlineData(true, ExcelWorksheetCopyMode.Values)]
        [InlineData(true, ExcelWorksheetCopyMode.Package)]
        public void Test_WorksheetCopy_PrintNamesCountTowardsDefinedNameLimits(bool merge, ExcelWorksheetCopyMode mode) {
            string sourcePath = Path.Combine(_directoryWithFiles, $"CopyNameBudgetSource-{merge}-{mode}.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, $"CopyNameBudgetTarget-{merge}-{mode}.xlsx");
            // The layout fixture copies four names: AllowedValues, PrintRows, the print area and the print titles.
            CreateLayoutSource(sourcePath);
            using (var createdTarget = ExcelDocument.Create(targetPath)) {
                createdTarget.AddSheet("Existing").Cell(1, 1).SetValue("Existing");
                createdTarget.Save();
            }

            using var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            using var targetDocument = ExcelDocument.Load(targetPath);
            void Copy(int maximumNames, int maximumCharacters) {
                if (merge) {
                    targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions {
                        CopyMode = mode,
                        TargetSheetNames = new Dictionary<string, string> { ["Data"] = "Copy" },
                        MaxDefinedNames = maximumNames,
                        MaxDefinedNameCharacters = maximumCharacters
                    });
                } else {
                    targetDocument.CopyWorksheetFrom(sourceDocument, "Data", "Copy", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions {
                        CopyMode = mode,
                        MaxDefinedNames = maximumNames,
                        MaxDefinedNameCharacters = maximumCharacters
                    });
                }
            }

            Assert.Throws<InvalidOperationException>(() => Copy(3, 10_000));
            Assert.Throws<InvalidOperationException>(() => Copy(10, 40));
            Assert.Equal(new[] { "Existing" }, targetDocument.Sheets.Select(sheet => sheet.Name));

            Copy(4, 10_000);
            Assert.Contains("Copy", targetDocument.Sheets.Select(sheet => sheet.Name));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_WorksheetValuesCopy_KeepsLayoutOfRowsWithoutIndexes(bool editedSource) {
            string sourcePath = Path.Combine(_directoryWithFiles, $"ImplicitRowsSource-{editedSource}.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, $"ImplicitRowsTarget-{editedSource}.xlsx");
            using (var createdSource = ExcelDocument.Create(sourcePath)) {
                createdSource.AddSheet("Data").Cell(1, 1).SetValue("Header");
                createdSource.Save();
            }

            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, true)) {
                Worksheet worksheet = GetWorksheetPartByNameForOperations(source, "Data").Worksheet!;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;
                // The first row takes its index from its cell, the second follows it, and neither cell names a column
                // the second time.
                sheetData.Append(
                    new Row(new Cell { CellReference = "A10", CellValue = new CellValue(5D) }) { Hidden = true, Height = 30D, CustomHeight = true },
                    new Row(new Cell { CellValue = new CellValue(6D) }) { Height = 40D, CustomHeight = true });
                worksheet.GetFirstChild<SheetDimension>()?.Remove();
                worksheet.Save();
            }

            using (var createdTarget = ExcelDocument.Create(targetPath)) {
                createdTarget.AddSheet("Existing").Cell(1, 1).SetValue("Existing");
                createdTarget.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions {
                AccessMode = editedSource ? OfficeIMO.DocumentAccessMode.ReadWrite : OfficeIMO.DocumentAccessMode.ReadOnly
            }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                if (editedSource) {
                    sourceDocument["Data"].Cell(1, 1).SetValue("Header");
                }

                targetDocument.CopyWorksheetFrom(sourceDocument, "Data", "Copy");
                targetDocument.Save();
            }

            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                Row[] rows = GetWorksheetPartByNameForOperations(target, "Copy").Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
                Row row10 = rows.Single(row => row.RowIndex!.Value == 10U);
                Assert.True(row10.Hidden!.Value);
                Assert.Equal(30D, row10.Height!.Value);
                Assert.Equal(40D, rows.Single(row => row.RowIndex!.Value == 11U).Height!.Value);
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            using (var reloaded = ExcelDocument.Load(targetPath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.True(reloaded["Copy"].TryGetCellText(10, 1, out string ten));
                Assert.Equal("5", ten);
                Assert.True(reloaded["Copy"].TryGetCellText(11, 1, out string eleven));
                Assert.Equal("6", eleven);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Theory]
        [InlineData(ExcelWorksheetCopyMode.Values)]
        [InlineData(ExcelWorksheetCopyMode.Package)]
        public void Test_WorksheetCopy_FormulaThresholdsKeepTheirNames(ExcelWorksheetCopyMode mode) {
            string sourcePath = Path.Combine(_directoryWithFiles, $"ThresholdNameSource-{mode}.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, $"ThresholdNameTarget-{mode}.xlsx");
            using (var createdSource = ExcelDocument.Create(sourcePath)) {
                ExcelSheet data = createdSource.AddSheet("Data");
                for (int row = 1; row <= 3; row++) data.Cell(row, 1).SetValue(row);
                data.Cell(1, 3).SetValue(2);
                createdSource.Save();
            }

            AddDefinedName(sourcePath, "Threshold", "Data!$C$1");
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, true)) {
                Worksheet worksheet = GetWorksheetPartByNameForOperations(source, "Data").Worksheet!;
                worksheet.InsertAfter(new ConditionalFormatting(new ConditionalFormattingRule(new ColorScale(
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min },
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Formula, Val = "Threshold" },
                    new Color { Rgb = "FFF8696B" },
                    new Color { Rgb = "FF63BE7B" })) { Type = ConditionalFormatValues.ColorScale, Priority = 1 }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A3" }
                }, worksheet.GetFirstChild<SheetData>()!);
                worksheet.Save();
            }

            using (var createdTarget = ExcelDocument.Create(targetPath)) {
                createdTarget.AddSheet("Existing").Cell(1, 1).SetValue("Existing");
                createdTarget.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Data", "Copy", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions { CopyMode = mode });
                targetDocument.Save();
            }

            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                Assert.Equal("Threshold", GetWorksheetPartByNameForOperations(target, "Copy").Worksheet!
                    .Descendants<ConditionalFormatValueObject>().Single(threshold => threshold.Type!.Value == ConditionalFormatValueObjectValues.Formula).Val!.Value);
                DefinedName threshold = target.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "Threshold");
                Assert.Equal("Copy!$C$1", threshold.Text.Replace("'", string.Empty));
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        private static void AddLinkedWorkbook(string path, string linkedFile) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            ExternalWorkbookPart externalPart = workbookPart.AddNewPart<ExternalWorkbookPart>();
            ExternalRelationship link = externalPart.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
                new Uri(linkedFile, UriKind.Relative));
            externalPart.ExternalLink = new ExternalLink(
                new ExternalBook(new SheetNames(new SheetName { Val = "Sheet1" })) { Id = link.Id });
            externalPart.ExternalLink.Save();
            workbookPart.Workbook!.ExternalReferences ??= new ExternalReferences();
            workbookPart.Workbook.ExternalReferences.Append(new ExternalReference { Id = workbookPart.GetIdOfPart(externalPart) });
            workbookPart.Workbook.Save();
        }
    }
}

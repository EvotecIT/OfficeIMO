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

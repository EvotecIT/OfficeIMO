using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelWorkbookMerge_ImportsSelectedSheetsWithPrefix() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.Source.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.Target.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddWorkSheet("North").CellValue(1, 1, "North value");
                source.AddWorkSheet("South").CellValue(1, 1, "South value");
                source.Save();
            }

            using (var target = ExcelDocument.Create(targetPath))
            using (var source = ExcelDocument.Load(sourcePath, readOnly: true)) {
                target.AddWorkSheet("Summary");
                ExcelWorkbookMergeResult result = target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                    SheetNames = new[] { "South" },
                    SheetNamePrefix = "Imported "
                });

                Assert.Equal(1, result.SheetCount);
                Assert.Equal(new[] { "South" }, result.SourceSheets);
                Assert.Equal(new[] { "Imported South" }, result.TargetSheets);
                Assert.True(target["Imported South"].TryGetCellText(1, 1, out var importedValue));
                Assert.Equal("South value", importedValue);
                target.Save();
            }

            using (var reloaded = ExcelDocument.Load(targetPath, readOnly: true)) {
                Assert.True(reloaded["Imported South"].TryGetCellText(1, 1, out var importedValue));
                Assert.Equal("South value", importedValue);
            }
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_StreamBackedWorkbookDoesNotForceSave() {
            using var targetStream = new MemoryStream();
            using var sourceStream = new MemoryStream();

            using (var source = ExcelDocument.Create(sourceStream, autoSave: false)) {
                source.AddWorkSheet("Source").CellValue(1, 1, "Imported");
                source.Save(sourceStream);
            }

            sourceStream.Position = 0;
            using (var target = ExcelDocument.Create(targetStream, autoSave: false))
            using (var source = ExcelDocument.Load(sourceStream, readOnly: true)) {
                target.AddWorkSheet("Target");
                ExcelWorkbookMergeResult result = target.MergeWorkbookFrom(source);

                Assert.Equal(1, result.SheetCount);
                Assert.True(target["Source"].TryGetCellText(1, 1, out var imported));
                Assert.Equal("Imported", imported);
                target.Save(targetStream);
            }

            targetStream.Position = 0;
            using var reloaded = ExcelDocument.Load(targetStream, readOnly: true);
            Assert.Equal(2, reloaded.Sheets.Count);
            Assert.True(reloaded["Source"].TryGetCellText(1, 1, out var value));
            Assert.Equal("Imported", value);
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_RewritesCopiedWorksheetFormulasForPrefixedNames() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.FormulaSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookMerge.FormulaTarget.xlsx");

            using (var source = ExcelDocument.Create(sourcePath)) {
                ExcelSheet data = source.AddWorkSheet("Data");
                data.CellValue(1, 1, 42);

                ExcelSheet summary = source.AddWorkSheet("Summary");
                summary.CellFormula(1, 1, "Data!A1");
                summary.ValidationCustomFormula("B2", "COUNTIF(Data!$A$1:$A$1,\">0\")>0");
                source.Save();
            }

            using (var target = ExcelDocument.Create(targetPath))
            using (var source = ExcelDocument.Load(sourcePath, readOnly: true)) {
                target.AddWorkSheet("Existing").CellValue(1, 1, "Existing");
                target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                    SheetNamePrefix = "Imported ",
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                target.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorksheetPart summaryPart = GetWorksheetPartByNameForOperations(spreadsheet, "Imported Summary");
                Cell formulaCell = summaryPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
                Formula1 validationFormula = Assert.Single(summaryPart.Worksheet.Descendants<Formula1>());

                Assert.Equal("'Imported Data'!A1", formulaCell.CellFormula?.Text);
                Assert.Equal("COUNTIF('Imported Data'!$A$1:$A$1,\">0\")>0", validationFormula.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }
    }
}

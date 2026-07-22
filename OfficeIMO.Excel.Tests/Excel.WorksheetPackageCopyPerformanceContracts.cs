using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeRewritesSharedStringsWithoutAdoptingSourceTable() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyIndexedSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyIndexedTarget.xlsx");

            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                ExcelSheet source = sourceDocument.AddWorksheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Revenue");
                source.CellValue(2, 1, "NA");
                source.CellValue(2, 2, 125.5m);
                source.CellBold(1, 1, true);
                source.CellBold(1, 2, true);
                source.AddTable("A1:B2", hasHeader: true, name: "SourceSales", OfficeIMO.Excel.TableStyle.TableStyleMedium9);
                sourceDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                ExcelSheet copied = targetDocument.CopyWorksheetFrom(
                    sourceDocument,
                    "Source",
                    "Imported",
                    SheetNameValidationMode.Sanitize,
                    new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package });
                targetDocument.Save();
                targetDocument.AddWorksheet("AfterCopy").CellValue(1, 1, "New shared string");
                targetDocument.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetPath, false)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                Assert.NotNull(workbookPart.SharedStringTablePart);
                Assert.Equal(
                    new[] { "New shared string" },
                    workbookPart.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>()
                        .Select(static item => item.InnerText));
                WorksheetPart importedPart = workbookPart.WorksheetParts.Single(part =>
                    part.TableDefinitionParts.Any(tablePart => tablePart.Table?.Name?.Value == "SourceSales"));
                Cell header = importedPart.Worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A1");
                Assert.Equal(CellValues.InlineString, header.DataType?.Value);
                Assert.Equal("Region", header.InnerText);
                Assert.NotNull(header.StyleIndex);
                Assert.NotEqual(0U, header.StyleIndex!.Value);
            }

            using (var targetDocument = ExcelDocument.Load(targetPath, new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                using var reader = targetDocument.CreateReader();
                object?[,] values = reader.GetSheet("Imported").ReadRange("A1:B2");
                Assert.Equal("Region", values[0, 0]);
                Assert.Equal("Revenue", values[0, 1]);
                Assert.Equal("NA", values[1, 0]);
                Assert.Equal(125.5d, values[1, 1]);
                Assert.Equal("New shared string", reader.GetSheet("AfterCopy").ReadRange("A1:A1")[0, 0]);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }
    }
}

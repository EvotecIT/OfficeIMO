using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FileBackedEdit_PersistsExplicitlyWithoutChangingLoadFastPath() {
            string path = Path.Combine(_directoryWithFiles, "FileBackedEdit.xlsx");
            using (var created = ExcelDocument.Create()) {
                created.AddWorksheet("Data").CellValue(1, 1, "before");
                created.Save(path);
            }

            using (ExcelDocument document = ExcelDocument.OpenFileBacked(path)) {
                Assert.True(document.UsesFileBackedPackage);
                document.Sheets[0].CellValue(1, 1, "after");
                document.Save();
                Assert.True(document.UsesFileBackedPackage);
            }

            using ExcelDocument loaded = ExcelDocument.Load(path);
            Assert.False(loaded.UsesFileBackedPackage);
            Assert.True(loaded.Sheets[0].TryGetCellValueSnapshot(1, 1, out ExcelCellValueSnapshot? value));
            Assert.Equal("after", value!.Text);
        }

        [Fact]
        public void Test_FileBackedEdit_EnforcesBudgetAndCancellationBeforeOpen() {
            string path = Path.Combine(_directoryWithFiles, "FileBackedBudget.xlsx");
            using (var created = ExcelDocument.Create()) {
                created.AddWorksheet("Data").CellValue(1, 1, "value");
                created.Save(path);
            }

            Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenFileBacked(path,
                new ExcelLoadOptions { MaxInputBytes = 1 }));

            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() =>
                ExcelDocument.OpenFileBacked(path, cancellationToken: cancellation.Token));
        }

        [Fact]
        public void Test_FileBackedEdit_SaveOnDisposeUsesAssociatedPath() {
            string path = Path.Combine(_directoryWithFiles, "FileBackedSaveOnDispose.xlsx");
            using (var created = ExcelDocument.Create()) {
                created.AddWorksheet("Data").CellValue(1, 1, 1);
                created.Save(path);
            }

            using (ExcelDocument document = ExcelDocument.OpenFileBacked(path, new ExcelLoadOptions {
                PersistenceMode = DocumentPersistenceMode.SaveOnDispose
            })) {
                document.Sheets[0].CellValue(1, 1, 2);
            }

            using ExcelDocument loaded = ExcelDocument.Load(path);
            Assert.True(loaded.Sheets[0].TryGetCellValueSnapshot(1, 1, out ExcelCellValueSnapshot? value));
            Assert.Equal("2", value!.Text);
        }

        [Fact]
        public void Test_FileBackedEdit_SaveUsesTemporaryBudgetWithoutManagedPackageLimit() {
            string path = Path.Combine(_directoryWithFiles, "FileBackedSaveBudget.xlsx");
            using (var created = ExcelDocument.Create()) {
                created.AddWorksheet("Data").CellValue(1, 1, "before");
                created.Save(path);
            }

            using (ExcelDocument document = ExcelDocument.OpenFileBacked(path)) {
                document.Sheets[0].CellValue(1, 1, "after");
                Assert.Throws<IOException>(() => document.Save(new ExcelSaveOptions {
                    MaxInMemoryPackageBytes = null,
                    MaxTemporaryPackageBytes = 1
                }));

                document.Save(new ExcelSaveOptions {
                    MaxInMemoryPackageBytes = 1,
                    MaxTemporaryPackageBytes = null,
                    ValidateOpenXml = true
                });
                Assert.True(document.UsesFileBackedPackage);
            }

            using ExcelDocument loaded = ExcelDocument.Load(path);
            Assert.Equal("after", loaded.Sheets[0].CellAt(1, 1).GetValue<string>());
        }

        [Fact]
        public void Test_FileBackedEdit_PreservesCarriageReturnsInLoadedWorksheet() {
            const string value = "First\r\nSecond\rThird\nFourth";
            string path = Path.Combine(_directoryWithFiles, "FileBackedCarriageReturns.xlsx");
            using (var created = ExcelDocument.Create()) {
                var sheet = created.AddWorksheet("Text");
                sheet.CellValue(1, 1, value);
                sheet.CellValue(1, 2, "inline");
                sheet.CellValue(1, 3, "plain");
                var cells = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                    .ToDictionary(cell => cell.CellReference!.Value!);
                cells["B1"].CellValue = null;
                cells["B1"].DataType = CellValues.InlineString;
                cells["B1"].InlineString = new InlineString(new Text(value));
                cells["C1"].DataType = CellValues.String;
                cells["C1"].CellValue = new CellValue(value);
                sheet.MarkRequiresSavePreparation();
                created.Save(path, new ExcelSaveOptions { DisableFastPackageWriter = true });
            }

            using (ExcelDocument document = ExcelDocument.OpenFileBacked(path)) {
                document.Sheets[0].CellValue(2, 1, "edited");
                document.Save(new ExcelSaveOptions { ValidateOpenXml = true });
            }

            using (var package = SpreadsheetDocument.Open(path, false)) {
                WorkbookPart workbook = package.WorkbookPart!;
                var writtenCells = workbook.WorksheetParts.Single().Worksheet
                    .Descendants<Cell>()
                    .ToDictionary(cell => cell.CellReference!.Value!);
                int sharedStringIndex = int.Parse(
                    writtenCells["A1"].CellValue!.Text,
                    CultureInfo.InvariantCulture);
                Assert.Equal(value, workbook.SharedStringTablePart!.SharedStringTable!
                    .Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText);
                Assert.Equal(value, writtenCells["B1"].InlineString!.InnerText);
                Assert.Equal(value, writtenCells["C1"].CellValue!.Text);
            }

            using var reopened = ExcelDocument.Load(path);
            for (int column = 1; column <= 3; column++) {
                Assert.True(reopened["Text"].TryGetCellText(1, column, out string actual));
                Assert.Equal(value, actual);
            }
        }
    }
}

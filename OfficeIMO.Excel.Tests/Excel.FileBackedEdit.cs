using System;
using System.IO;
using System.Threading;
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
    }
}

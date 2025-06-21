using System;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_SimpleExcelDocumentCreation() {
            var filePath = Path.Combine(_directoryWithFiles, "TestFileTemporary3.xlsx");

            var path = File.Exists(filePath);
            File.Delete(filePath);

            Assert.False(path); // MUST BE FALSE

            using (var document = ExcelDocument.Create(filePath)) {
                document.Save();

                path = File.Exists(filePath);
                Assert.True(path);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_CreatingExcel1() {
            var filePath = Path.Combine(_directoryWithFiles, "TestFileTemporary1.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                Assert.True(document.Sheets.Count == 0);
                var sheet1 = document.AddWorkSheet("Test1");
                var sheet2 = document.AddWorkSheet("Test2");
                var sheet3 = document.AddWorkSheet("Test3");

                Assert.True(document.Sheets.Count == 3);
                Assert.True(document.Sheets[0].Name == "Test1");
                Assert.True(document.Sheets[1].Name == "Test2");
                Assert.True(document.Sheets[2].Name == "Test3");
                document.Save(false);
            }
        }


        [Fact]
        public void Test_CreatingExcel2() {
            var filePath = Path.Combine(_directoryWithFiles, "TestFileTemporary2.xlsx");
            using (var document = ExcelDocument.Create(filePath, "WorkSheet5")) {
                Assert.True(document.Sheets.Count == 1);
                ExcelSheet sheet = document.AddWorkSheet("Test");
                Assert.True(document.Sheets.Count == 2);
                Assert.True(document.Sheets[0].Name == "WorkSheet5");
                Assert.True(document.Sheets[1].Name == "Test");
                document.Save(false);
            }
        }

        [Fact]
        public void Test_OpeningExcel() {
            using (var document = ExcelDocument.Load(Path.Combine(_directoryDocuments, "BasicExcel.xlsx"))) {
                Assert.True(document.FilePath != null);
                Assert.True(document.Sheets.Count == 4);

                Assert.True(document.Sheets[0].Name == "Sheet1");
                Assert.True(document.Sheets[1].Name == "Sheet2");
                Assert.True(document.Sheets[2].Name == "DifferentTest");
                Assert.True(document.Sheets[3].Name == "Sheet3");
                document.Save();
            }
        }
    }
}

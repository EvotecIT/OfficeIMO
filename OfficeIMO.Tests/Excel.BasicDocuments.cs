using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_SimpleExcelDocumentCreation() {
            var filePath = Path.Combine(_directoryWithFiles, "TestFileTemporary.xlsx");



            var path = File.Exists(filePath);
            File.Delete(filePath);

            Assert.False(path); // MUST BE FALSE

            ExcelDocument document = ExcelDocument.Create(filePath);

            document.Save();

            path = File.Exists(filePath);
            Assert.True(path);
            document.Dispose();

            File.Delete(filePath);
        }

        [Fact]
        public void Test_OpeningExcel() {
            using (ExcelDocument document = ExcelDocument.Load(Path.Combine(_directoryWithFiles, "BasicExcel.xlsx"))) {

                Assert.True(document.FilePath != null);
                Assert.True(document.Sheets.Count == 4);
                
                document.Save();
            }
        }
    }
}
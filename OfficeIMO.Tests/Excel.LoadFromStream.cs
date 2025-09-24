using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public partial class Excel
    {
        [Fact]
        public void Load_FromMemoryStream_PreservesSharedStringsAndProperties()
        {
            string filePath = Path.Combine(_directoryWithFiles, "LoadFromStream.xlsx");

            try
            {
                using (var document = ExcelDocument.Create(filePath))
                {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Alpha");
                    sheet.CellValue(2, 1, "Beta");
                    sheet.CellValue(3, 1, "Gamma");

                    document.BuiltinDocumentProperties.Title = "Stream Title";
                    document.BuiltinDocumentProperties.Creator = "Stream Creator";
                    document.ApplicationProperties.Company = "Stream Company";
                    document.ApplicationProperties.Manager = "Stream Manager";

                    document.Save();
                }

                string[] expectedShared;
                using (var fromFile = ExcelDocument.Load(filePath))
                {
                    expectedShared = fromFile._spreadSheetDocument.WorkbookPart!
                        .SharedStringTablePart!
                        .SharedStringTable!
                        .Elements<SharedStringItem>()
                        .Select(item => item.InnerText)
                        .ToArray();
                }

                using var memory = new MemoryStream(File.ReadAllBytes(filePath));
                using var fromStream = ExcelDocument.Load(memory);

                Assert.Equal("Stream Title", fromStream.BuiltinDocumentProperties.Title);
                Assert.Equal("Stream Creator", fromStream.BuiltinDocumentProperties.Creator);
                Assert.Equal("Stream Company", fromStream.ApplicationProperties.Company);
                Assert.Equal("Stream Manager", fromStream.ApplicationProperties.Manager);

                var actualShared = fromStream._spreadSheetDocument.WorkbookPart!
                    .SharedStringTablePart!
                    .SharedStringTable!
                    .Elements<SharedStringItem>()
                    .Select(item => item.InnerText)
                    .ToArray();

                Assert.Equal(expectedShared, actualShared);
            }
            finally
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task LoadAsync_FromMemoryStream_PreservesSharedStringsAndProperties()
        {
            string filePath = Path.Combine(_directoryWithFiles, "LoadFromStreamAsync.xlsx");

            try
            {
                using (var document = ExcelDocument.Create(filePath))
                {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Delta");
                    sheet.CellValue(2, 1, "Epsilon");
                    sheet.CellValue(3, 1, "Zeta");

                    document.BuiltinDocumentProperties.Title = "Async Stream Title";
                    document.BuiltinDocumentProperties.Creator = "Async Stream Creator";
                    document.ApplicationProperties.Company = "Async Stream Company";
                    document.ApplicationProperties.Manager = "Async Stream Manager";

                    document.Save();
                }

                string[] expectedShared;
                using (var fromFile = ExcelDocument.Load(filePath))
                {
                    expectedShared = fromFile._spreadSheetDocument.WorkbookPart!
                        .SharedStringTablePart!
                        .SharedStringTable!
                        .Elements<SharedStringItem>()
                        .Select(item => item.InnerText)
                        .ToArray();
                }

                await using var memory = new MemoryStream(File.ReadAllBytes(filePath));
                await using var fromStream = await ExcelDocument.LoadAsync(memory);

                Assert.Equal("Async Stream Title", fromStream.BuiltinDocumentProperties.Title);
                Assert.Equal("Async Stream Creator", fromStream.BuiltinDocumentProperties.Creator);
                Assert.Equal("Async Stream Company", fromStream.ApplicationProperties.Company);
                Assert.Equal("Async Stream Manager", fromStream.ApplicationProperties.Manager);

                var actualShared = fromStream._spreadSheetDocument.WorkbookPart!
                    .SharedStringTablePart!
                    .SharedStringTable!
                    .Elements<SharedStringItem>()
                    .Select(item => item.InnerText)
                    .ToArray();

                Assert.Equal(expectedShared, actualShared);
            }
            finally
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
        }
    }
}


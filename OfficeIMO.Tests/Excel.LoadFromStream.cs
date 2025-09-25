using System;
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

                using var memory = new MemoryStream(File.ReadAllBytes(filePath));
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

        [Fact]
        public void Load_FromMemoryStream_WithAutoSave_PersistsChanges()
        {
            string filePath = Path.Combine(_directoryWithFiles, "LoadFromStreamAutoSave.xlsx");

            try
            {
                using (var document = ExcelDocument.Create(filePath))
                {
                    var sheet = document.AddWorkSheet("AutoSave");
                    sheet.CellValue(1, 1, "Original");
                    document.Save();
                }

                var bytes = File.ReadAllBytes(filePath);
                using var memory = new MemoryStream();
                memory.Write(bytes, 0, bytes.Length);
                memory.Seek(0, SeekOrigin.Begin);

                using (var document = ExcelDocument.Load(memory, readOnly: false, autoSave: true))
                {
                    var sheet = document.Sheets[0];
                    sheet.CellValue(1, 1, "Updated");
                }

                memory.Seek(0, SeekOrigin.Begin);
                using var reloaded = ExcelDocument.Load(memory);
                var reloadedSheet = reloaded.Sheets[0];
                Assert.True(reloadedSheet.TryGetCellText(1, 1, out var text));
                Assert.Equal("Updated", text);
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
        public async Task LoadAsync_FromMemoryStream_WithAutoSave_PersistsChanges()
        {
            string filePath = Path.Combine(_directoryWithFiles, "LoadFromStreamAutoSaveAsync.xlsx");

            try
            {
                using (var document = ExcelDocument.Create(filePath))
                {
                    var sheet = document.AddWorkSheet("AutoSave");
                    sheet.CellValue(1, 1, "Original Async");
                    document.Save();
                }

                var bytes = File.ReadAllBytes(filePath);
                using var memory = new MemoryStream();
                memory.Write(bytes, 0, bytes.Length);
                memory.Seek(0, SeekOrigin.Begin);

                await using (var document = await ExcelDocument.LoadAsync(memory, readOnly: false, autoSave: true))
                {
                    var sheet = document.Sheets[0];
                    sheet.CellValue(1, 1, "Updated Async");
                }

                memory.Seek(0, SeekOrigin.Begin);
                using var reloaded = ExcelDocument.Load(memory);
                var reloadedSheet = reloaded.Sheets[0];
                Assert.True(reloadedSheet.TryGetCellText(1, 1, out var text));
                Assert.Equal("Updated Async", text);
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
        public void Load_FromReadOnlyStream_WithAutoSave_Throws()
        {
            string filePath = Path.Combine(_directoryWithFiles, "LoadFromStreamReadOnly.xlsx");

            try
            {
                using (var document = ExcelDocument.Create(filePath))
                {
                    document.AddWorkSheet("Readonly");
                    document.Save();
                }

                using var readOnlyStream = new MemoryStream(File.ReadAllBytes(filePath), writable: false);
                var ex = Assert.Throws<ArgumentException>(() => ExcelDocument.Load(readOnlyStream, readOnly: false, autoSave: true));
                Assert.Equal("stream", ex.ParamName);
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
        public async Task LoadAsync_FromReadOnlyStream_WithAutoSave_Throws()
        {
            string filePath = Path.Combine(_directoryWithFiles, "LoadFromStreamReadOnlyAsync.xlsx");

            try
            {
                using (var document = ExcelDocument.Create(filePath))
                {
                    document.AddWorkSheet("Readonly");
                    document.Save();
                }

                using var readOnlyStream = new MemoryStream(File.ReadAllBytes(filePath), writable: false);
                var ex = await Assert.ThrowsAsync<ArgumentException>(() => ExcelDocument.LoadAsync(readOnlyStream, readOnly: false, autoSave: true));
                Assert.Equal("stream", ex.ParamName);
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


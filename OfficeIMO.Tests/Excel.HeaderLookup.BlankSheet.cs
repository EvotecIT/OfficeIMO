using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public partial class Excel
    {
        [Fact]
        public void Test_BlankSheetHeaderLookups()
        {
            var filePath = Path.Combine(_directoryWithFiles, "BlankSheetHeaderLookups.xlsx");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            using (var document = ExcelDocument.Create(filePath))
            {
                document.AddWorkSheet("Empty");
                document.Save();
            }

            try
            {
                using var document = ExcelDocument.Load(filePath);
                var sheet = document.Sheets[0];

                var headers = sheet.GetHeaderMap();
                Assert.Empty(headers);

                Assert.False(sheet.TryGetColumnIndexByHeader("Missing", out var columnIndex));
                Assert.Equal(0, columnIndex);

                Assert.False(sheet.TryGetColumnIndexByHeader("Column1", out _));

                Assert.Null(Record.Exception(() => sheet.SetByHeader(2, "Missing", "Value")));
                Assert.Null(Record.Exception(() => sheet.LinkByHeaderToInternalSheets("Missing")));
                Assert.False(sheet.TryLinkByHeaderToInternalSheets("Missing"));
                Assert.Null(Record.Exception(() => sheet.AutoFilterByHeaderEquals("Missing", new[] { "v" })));
                Assert.Null(Record.Exception(() => sheet.AutoFilterByHeaderContains("Missing", "v")));
                Assert.Null(Record.Exception(() => sheet.AutoFilterByHeadersEquals(("Missing", new[] { "v" }))));
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


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

                Assert.False(sheet.TryGetColumnIndexByHeader("Column1", out var column1Index));
                Assert.Equal(0, column1Index);

                var setEx = Record.Exception(() => sheet.SetByHeader(2, "Missing", "Value"));
                Assert.Null(setEx);

                var styleEx = Record.Exception(() => sheet.ColumnStyleByHeader("Missing"));
                Assert.Null(styleEx);

                var linkInternalEx = Record.Exception(() => sheet.LinkByHeaderToInternalSheets("Missing"));
                Assert.Null(linkInternalEx);

                var linkExternalEx = Record.Exception(() => sheet.LinkByHeaderToUrls("Missing", urlForCellText: text => text));
                Assert.Null(linkExternalEx);

                Assert.False(sheet.TryLinkByHeaderToInternalSheets("Missing"));
                Assert.False(sheet.TryLinkByHeaderToUrls("Missing", urlForCellText: _ => "https://example.com"));
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

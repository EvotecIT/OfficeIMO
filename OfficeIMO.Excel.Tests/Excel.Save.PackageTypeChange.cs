using OfficeIMO.Excel;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(".xltx", false, false)]
        [InlineData(".xlsm", false, false)]
        [InlineData(".xltx", true, false)]
        [InlineData(".xlsm", true, false)]
        [InlineData(".xltx", false, true)]
        [InlineData(".xlsm", false, true)]
        [InlineData(".xltx", true, true)]
        [InlineData(".xlsm", true, true)]
        public void Save_as_another_package_type_preserves_retained_worksheet_edits(string extension, bool standardWriter, bool previouslySaved) {
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-TypeChange-" + Guid.NewGuid().ToString("N") + extension);
            try {
                using var document = ExcelDocument.Create();
                document.SheetCachingEnabled = false;
                ExcelSheet sheet = document.AddWorksheet("Report");
                ExcelSheet alias = document.Sheets[0];
                sheet.CellValue(1, 1, "Before format change");
                var options = new ExcelSaveOptions { DisableFastPackageWriter = standardWriter };
                if (previouslySaved) {
                    using var initial = new MemoryStream();
                    document.Save(initial, options);
                }
                document.Save(path, options);
                sheet.CellValue(2, 1, "After format change");
                alias.CellWrapText(2, 1);
                document.Save(path, options);

                using var saved = ExcelDocument.Load(path);
                Assert.Equal("Before format change", saved["Report"].CellAt(1, 1).GetValue<string>());
                Assert.Equal("After format change", saved["Report"].CellAt(2, 1).GetValue<string>());
                Assert.True(saved["Report"].CellAt(2, 1).GetStyle().WrapText);
                Assert.Empty(sheet.GetTables());
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Promoting_to_a_macro_workbook_preserves_retained_worksheet_edits(bool nativeProject) {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Report");
            document.AddWorksheet("Keep");
            sheet.CellValue(1, 1, "Before promotion");
            var project = OfficeVbaProject.Create();
            project.AddModule("Helpers", "' Inert worksheet lifetime fixture\r\n");
            if (nativeProject) document.SetVbaProject(project);
            else document.AddMacro(project.Write().GetBytes());
            Assert.Empty(sheet.GetTables());
            sheet.CellValue(2, 1, "After promotion");
            sheet.CellWrapText(2, 1);
            sheet.SetHidden(true);
            document.Save(stream);

            using var saved = ExcelDocument.Load(stream);
            Assert.True(saved.HasMacros);
            Assert.Equal("Before promotion", saved["Report"].CellAt(1, 1).GetValue<string>());
            Assert.Equal("After promotion", saved["Report"].CellAt(2, 1).GetValue<string>());
            Assert.True(saved["Report"].CellAt(2, 1).GetStyle().WrapText);
            Assert.True(saved["Report"].Hidden);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task Async_and_encrypted_template_saves_preserve_retained_worksheet_edits(bool encrypted) {
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-TypeChange-" + Guid.NewGuid().ToString("N") + ".xltx");
            try {
                using var document = ExcelDocument.Create();
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Before conversion");
                if (encrypted) document.SaveEncrypted(path, "inert-test-password");
                else await document.SaveAsync(path);
                sheet.CellValue(2, 1, "After conversion");
                sheet.CellWrapText(2, 1);
                if (encrypted) document.SaveEncrypted(path, "inert-test-password");
                else await document.SaveAsync(path);

                using var saved = encrypted ? ExcelDocument.LoadEncrypted(path, "inert-test-password") : ExcelDocument.Load(path);
                Assert.Equal("Before conversion", saved["Report"].CellAt(1, 1).GetValue<string>());
                Assert.Equal("After conversion", saved["Report"].CellAt(2, 1).GetValue<string>());
                Assert.True(saved["Report"].CellAt(2, 1).GetStyle().WrapText);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}

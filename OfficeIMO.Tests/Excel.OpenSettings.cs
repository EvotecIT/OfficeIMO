using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelLoad_CustomOpenSettingsAreApplied() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelCustomOpenSettings.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = ExcelDocument.Create(filePath)) {
                document.Save();
            }

            var settings = new OpenSettings {
                MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2016)
            };

            using (var document = ExcelDocument.Load(filePath, openSettings: settings)) {
                var mcSettings = document._spreadSheetDocument.MarkupCompatibilityProcessSettings;
                Assert.NotNull(mcSettings);
                Assert.Equal(MarkupCompatibilityProcessMode.ProcessAllParts, mcSettings.ProcessMode);
                Assert.Equal(FileFormatVersions.Office2016, mcSettings.TargetFileFormatVersions);
            }

            File.Delete(filePath);
        }

        [Fact]
        public async Task Test_ExcelLoadAsync_CustomOpenSettingsAreApplied() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelCustomOpenSettingsAsync.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = ExcelDocument.Create(filePath)) {
                document.Save();
            }

            var settings = new OpenSettings {
                MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2007)
            };

            await using (var document = await ExcelDocument.LoadAsync(filePath, openSettings: settings)) {
                var mcSettings = document._spreadSheetDocument.MarkupCompatibilityProcessSettings;
                Assert.NotNull(mcSettings);
                Assert.Equal(MarkupCompatibilityProcessMode.ProcessAllParts, mcSettings.ProcessMode);
                Assert.Equal(FileFormatVersions.Office2007, mcSettings.TargetFileFormatVersions);
            }

            File.Delete(filePath);
        }
    }
}

using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_WordLoad_CustomOpenSettingsAreApplied() {
            var filePath = Path.Combine(_directoryWithFiles, "WordCustomOpenSettings.docx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = WordDocument.Create(filePath)) {
                document.AddParagraph("Custom settings");
                document.Save();
            }

            var settings = new OpenSettings {
                MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2010)
            };

            using (var document = WordDocument.Load(filePath, openSettings: settings)) {
                var mcSettings = document._wordprocessingDocument.MarkupCompatibilityProcessSettings;
                Assert.NotNull(mcSettings);
                Assert.Equal(MarkupCompatibilityProcessMode.ProcessAllParts, mcSettings.ProcessMode);
                Assert.Equal(FileFormatVersions.Office2010, mcSettings.TargetFileFormatVersions);
            }

            File.Delete(filePath);
        }

        [Fact]
        public async Task Test_WordLoadAsync_CustomOpenSettingsAreApplied() {
            var filePath = Path.Combine(_directoryWithFiles, "WordCustomOpenSettingsAsync.docx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = WordDocument.Create(filePath)) {
                document.AddParagraph("Custom settings async");
                document.Save();
            }

            var settings = new OpenSettings {
                MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013)
            };

            await using (var document = await WordDocument.LoadAsync(filePath, openSettings: settings, cancellationToken: CancellationToken.None)) {
                var mcSettings = document._wordprocessingDocument.MarkupCompatibilityProcessSettings;
                Assert.NotNull(mcSettings);
                Assert.Equal(MarkupCompatibilityProcessMode.ProcessAllParts, mcSettings.ProcessMode);
                Assert.Equal(FileFormatVersions.Office2013, mcSettings.TargetFileFormatVersions);
            }

            File.Delete(filePath);
        }
    }
}

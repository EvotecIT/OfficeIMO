using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_PageSizeAndOrientationSwitch() {
            string filePath = Path.Combine(_directoryWithFiles, "PageSizeOrientationSwitch.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].PageSettings.PageSize = WordPageSize.A4;
                document.Sections[0].PageSettings.Orientation = PageOrientationValues.Landscape;
                document.AddParagraph("Test");
                document.AddSection();
                document.Sections[1].PageSettings.PageSize = WordPageSize.A5;
                document.Sections[1].PageSettings.Orientation = PageOrientationValues.Portrait;
                document.AddParagraph("Section 1");
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(PageOrientationValues.Landscape, document.Sections[0].PageSettings.Orientation);
                Assert.Equal(WordPageSize.A4, document.Sections[0].PageSettings.PageSize);
                Assert.Equal(PageOrientationValues.Portrait, document.Sections[1].PageSettings.Orientation);
                Assert.Equal(WordPageSize.A5, document.Sections[1].PageSettings.PageSize);
            }
        }
    }
}

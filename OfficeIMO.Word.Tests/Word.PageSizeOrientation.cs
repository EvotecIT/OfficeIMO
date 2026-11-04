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
                document.Sections[0].PageSetup.Size.PageSize = WordPageSize.A4;
                document.Sections[0].PageSetup.Size.Orientation = OfficePageOrientation.Landscape;
                document.AddParagraph("Test");
                document.AddSection();
                document.Sections[1].PageSetup.Size.PageSize = WordPageSize.A5;
                document.Sections[1].PageSetup.Size.Orientation = OfficePageOrientation.Portrait;
                document.AddParagraph("Section 1");
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(OfficePageOrientation.Landscape, document.Sections[0].PageSetup.Size.Orientation);
                Assert.Equal(WordPageSize.A4, document.Sections[0].PageSetup.Size.PageSize);
                Assert.Equal(OfficePageOrientation.Portrait, document.Sections[1].PageSetup.Size.Orientation);
                Assert.Equal(WordPageSize.A5, document.Sections[1].PageSetup.Size.PageSize);
            }
        }
    }
}

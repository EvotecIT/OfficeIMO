using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_WordDocumentStatisticsCounts() {
            var filePath = Path.Combine(_directoryWithFiles, "DocumentStatistics.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AddParagraph("This is first paragraph");
                document.AddParagraph("Second paragraph with image");
                document.Paragraphs[1].AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));

                Assert.Equal(2, document.Statistics.Paragraphs);
                Assert.Equal(1, document.Statistics.Images);
                Assert.Equal(1, document.Statistics.Pages);
                Assert.True(document.Statistics.Words >= 8);

                Assert.Equal(0, document.Statistics.Tables);
                Assert.Equal(0, document.Statistics.Charts);
                Assert.Equal(0, document.Statistics.Shapes);
                Assert.Equal(0, document.Statistics.Bookmarks);
                Assert.Equal(0, document.Statistics.Lists);
                Assert.True(document.Statistics.CharactersWithSpaces >= 50);
                Assert.True(document.Statistics.Characters >= 40);

                document.Save(false);
            }
        }
    }
}

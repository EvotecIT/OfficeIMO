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
                Assert.Equal(8, document.Statistics.Words);

                Assert.Equal(0, document.Statistics.Tables);
                Assert.Equal(0, document.Statistics.Charts);
                Assert.Equal(0, document.Statistics.Shapes);
                Assert.Equal(0, document.Statistics.Bookmarks);
                Assert.Equal(0, document.Statistics.Lists);
                Assert.Equal(50, document.Statistics.CharactersWithSpaces);
                Assert.Equal(44, document.Statistics.Characters);

                WordDocumentStatistics statistics = document.Statistics;
                WordDocumentStatisticsInfo initial = document.Inspect(includeImageBytes: false).Statistics;
                Assert.Same(statistics, document.Statistics);
                Assert.Equal(1, initial.Pages);
                Assert.Equal(2, initial.Paragraphs);
                Assert.Equal(8, initial.Words);
                Assert.Equal(1, initial.Images);
                Assert.Equal(0, initial.Tables);
                Assert.Equal(0, initial.Charts);
                Assert.Equal(0, initial.Shapes);
                Assert.Equal(0, initial.Bookmarks);
                Assert.Equal(0, initial.Lists);
                Assert.Equal(50, initial.CharactersWithSpaces);
                Assert.Equal(44, initial.Characters);

                document.AddParagraph("Third");
                WordDocumentStatisticsInfo updated = document.Inspect(includeImageBytes: false).Statistics;
                Assert.Same(statistics, document.Statistics);
                Assert.Equal(3, statistics.Paragraphs);
                Assert.Equal(9, statistics.Words);
                Assert.Equal(55, statistics.CharactersWithSpaces);
                Assert.Equal(49, statistics.Characters);
                Assert.Equal(3, updated.Paragraphs);
                Assert.Equal(9, updated.Words);
                Assert.Equal(55, updated.CharactersWithSpaces);
                Assert.Equal(49, updated.Characters);
                Assert.Equal(2, initial.Paragraphs);
                Assert.Equal(8, initial.Words);
                Assert.Equal(50, initial.CharactersWithSpaces);
                Assert.Equal(44, initial.Characters);

                document.ApplicationProperties.Pages = "0";
                document.ApplicationProperties.Words = 0;
                document.ApplicationProperties.Characters = "0";
                document.ApplicationProperties.CharactersWithSpaces = "0";
                WordDocumentStatisticsInfo zero = document.Inspect(includeImageBytes: false).Statistics;
                Assert.Equal(0, zero.Pages);
                Assert.Equal(0, zero.Words);
                Assert.Equal(0, zero.Characters);
                Assert.Equal(0, zero.CharactersWithSpaces);
                Assert.Equal(3, zero.Paragraphs);

                document.ApplicationProperties.Pages = "7";
                document.ApplicationProperties.Words = 42;
                document.ApplicationProperties.Characters = "1234";
                document.ApplicationProperties.CharactersWithSpaces = "1500";
                WordDocumentStatisticsInfo savedCounts = document.Inspect(includeImageBytes: false).Statistics;
                Assert.Equal(7, savedCounts.Pages);
                Assert.Equal(42, savedCounts.Words);
                Assert.Equal(1234, savedCounts.Characters);
                Assert.Equal(1500, savedCounts.CharactersWithSpaces);
                Assert.Equal(3, savedCounts.Paragraphs);
                Assert.Same(statistics, document.Statistics);

                document.ApplicationProperties.Pages = string.Empty;
                document.ApplicationProperties.Words = null;
                document.ApplicationProperties.Characters = string.Empty;
                document.ApplicationProperties.CharactersWithSpaces = string.Empty;
                Assert.Equal(1, statistics.Pages);
                Assert.Equal(9, statistics.Words);
                Assert.Equal(55, statistics.CharactersWithSpaces);
                Assert.Equal(49, statistics.Characters);

                document.AddPageBreak();
                Assert.Equal(2, statistics.Pages);
                Assert.Equal(2, document.Inspect(includeImageBytes: false).Statistics.Pages);

                document.Save();
            }
        }
    }
}

using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeDocumentWordStatisticsTests {
    [Fact]
    public void DetailedWordStatisticsPreserveAllNativeCountersAndDetachFromPathAndUnsavedModels() {
        string root = Directory.CreateTempSubdirectory("officeimo-word-statistics-").FullName;
        try {
            string path = Path.Combine(root, "source.docx"), moved = Path.Combine(root, "moved.docx");
            string picture = Path.Combine(root, "pixel.png");
            File.WriteAllBytes(picture, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII="));
            using (var source = WordDocument.Create()) {
                source.AddParagraph("Native statistics evidence").AddBookmark("Evidence");
                source.AddParagraph().AddImage(picture);
                source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Cell evidence";
                var chart = source.AddChart("Statistics chart");
                chart.AddCategories(new[] { "Category" });
                chart.AddBar("Series", new[] { 1 }, OfficeColor.Blue);
                source.AddShape(WordShapeType.Rectangle, OfficeLength.Points(20), OfficeLength.Points(10));
                source.AddList(WordListStyle.Numbered).AddItem("List evidence");
                source.Save(path);
            }
            int[] expected;
            using (var source = WordDocument.Load(path, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly }))
                expected = Values(source.Statistics);
            Assert.All(expected, value => Assert.True(value > 0));
            byte[] original = File.ReadAllBytes(path);
            var settings = new OfficeInspectOptions { IncludeSummary = true, IncludeContent = false, Detailed = true };
            var pathResult = OfficeDocument.Inspect(path, settings);
            var pathStatistics = Assert.IsType<WordDocumentInfo>(pathResult.Details).Statistics;
            Assert.Equal(expected, Values(pathStatistics));
            Assert.Equal(expected[0], pathResult.PageCount);
            Assert.Equal(original, File.ReadAllBytes(path));
            File.Move(path, moved);
            WordDocumentStatisticsInfo liveStatistics;
            int[] liveExpected;
            using (var live = WordDocument.Load(moved, new WordLoadOptions { PersistenceMode = DocumentPersistenceMode.Explicit })) {
                live.AddParagraph("Unsaved statistics remain visible");
                live.AddTable(1, 1);
                liveExpected = Values(live.Statistics);
                Assert.True(liveExpected[1] > expected[1]);
                Assert.True(liveExpected[4] > expected[4]);
                File.Delete(moved);
                var liveResult = OfficeDocument.Inspect((object)live, settings);
                liveStatistics = Assert.IsType<WordDocumentInfo>(liveResult.Details).Statistics;
                Assert.Equal(liveExpected, Values(liveStatistics));
                Assert.Equal(liveExpected, Values(live.Statistics));
                Assert.Contains(live.Paragraphs, paragraph => paragraph.Text == "Unsaved statistics remain visible");
                live.AddParagraph("Caller still owns the model");
                Assert.Equal(liveExpected, Values(liveStatistics));
                Assert.Equal(expected, Values(pathStatistics));
                Assert.False(File.Exists(moved));
            }
            Assert.Equal(liveExpected, Values(liveStatistics));
            Assert.Equal(expected, Values(pathStatistics));
        } finally { Directory.Delete(root, true); }
    }

    private static int[] Values(WordDocumentStatistics statistics) => new[] {
        statistics.Pages, statistics.Paragraphs, statistics.Words, statistics.Images, statistics.Tables,
        statistics.Charts, statistics.Shapes, statistics.Bookmarks, statistics.Lists,
        statistics.Characters, statistics.CharactersWithSpaces
    };

    private static int[] Values(WordDocumentStatisticsInfo statistics) => new[] {
        statistics.Pages, statistics.Paragraphs, statistics.Words, statistics.Images, statistics.Tables,
        statistics.Charts, statistics.Shapes, statistics.Bookmarks, statistics.Lists,
        statistics.Characters, statistics.CharactersWithSpaces
    };
}

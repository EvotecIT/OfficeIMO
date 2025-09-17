using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CleanupDocument_RemovesClutterFromHeadersAndFooters() {
            string filePath = Path.Combine(_directoryWithFiles, "CleanupHeadersAndFooters.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();

                var headerParagraph = document.Header!.Default.AddParagraph("Header ");
                headerParagraph.AddText("clutter ");
                headerParagraph.AddText("text");
                document.Header!.Default.AddParagraph();

                var footerParagraph = document.Footer!.Default.AddParagraph("Footer ");
                footerParagraph.AddText("clutter ");
                footerParagraph.AddText("text");
                document.Footer!.Default.AddParagraph();

                Assert.True(document.Header!.Default.Paragraphs.Count > 1);
                Assert.True(headerParagraph.GetRuns().Count() > 1);
                Assert.True(document.Footer!.Default.Paragraphs.Count > 1);
                Assert.True(footerParagraph.GetRuns().Count() > 1);

                document.CleanupDocument();

                Assert.True(document.Header!.Default.Paragraphs.Count == 1);
                Assert.True(document.Header!.Default.Paragraphs[0].GetRuns().Count() == 1);
                Assert.True(document.Header!.Default.Paragraphs[0].Text == "Header clutter text");

                Assert.True(document.Footer!.Default.Paragraphs.Count == 1);
                Assert.True(document.Footer!.Default.Paragraphs[0].GetRuns().Count() == 1);
                Assert.True(document.Footer!.Default.Paragraphs[0].Text == "Footer clutter text");

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CleanupHeadersAndFooters.docx"))) {
                Assert.True(document.Header!.Default.Paragraphs.Count == 1);
                Assert.True(document.Header!.Default.Paragraphs[0].GetRuns().Count() == 1);
                Assert.True(document.Header!.Default.Paragraphs[0].Text == "Header clutter text");

                Assert.True(document.Footer!.Default.Paragraphs.Count == 1);
                Assert.True(document.Footer!.Default.Paragraphs[0].GetRuns().Count() == 1);
                Assert.True(document.Footer!.Default.Paragraphs[0].Text == "Footer clutter text");
            }
        }
    }
}

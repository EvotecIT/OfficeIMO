using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CleanupDocument_RemovesClutterFromHeadersAndFooters() {
            string filePath = Path.Combine(_directoryWithFiles, "CleanupHeadersAndFooters.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();

                var defaultHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                var headerParagraph = defaultHeader.AddParagraph("Header ");
                headerParagraph.AddText("clutter ");
                headerParagraph.AddText("text");
                defaultHeader.AddParagraph();

                var defaultFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                var footerParagraph = defaultFooter.AddParagraph("Footer ");
                footerParagraph.AddText("clutter ");
                footerParagraph.AddText("text");
                defaultFooter.AddParagraph();

                Assert.True(defaultHeader.Paragraphs.Count > 1);
                Assert.True(headerParagraph.GetRuns().Count() > 1);
                Assert.True(defaultFooter.Paragraphs.Count > 1);
                Assert.True(footerParagraph.GetRuns().Count() > 1);

                document.CleanupDocument();

                Assert.True(defaultHeader.Paragraphs.Count == 1);
                Assert.True(defaultHeader.Paragraphs[0].GetRuns().Count() == 1);
                Assert.True(defaultHeader.Paragraphs[0].Text == "Header clutter text");

                Assert.True(defaultFooter.Paragraphs.Count == 1);
                Assert.True(defaultFooter.Paragraphs[0].GetRuns().Count() == 1);
                Assert.True(defaultFooter.Paragraphs[0].Text == "Footer clutter text");

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CleanupHeadersAndFooters.docx"))) {
                var reloadedHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                Assert.True(reloadedHeader.Paragraphs.Count == 1);
                Assert.True(reloadedHeader.Paragraphs[0].GetRuns().Count() == 1);
                Assert.True(reloadedHeader.Paragraphs[0].Text == "Header clutter text");

                var reloadedFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                Assert.True(reloadedFooter.Paragraphs.Count == 1);
                Assert.True(reloadedFooter.Paragraphs[0].GetRuns().Count() == 1);
                Assert.True(reloadedFooter.Paragraphs[0].Text == "Footer clutter text");
            }
        }
    }
}

using System;
using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void FormattingHelper_GetsRunsWithFlags() {
            using MemoryStream ms = new MemoryStream();
            using (var document = WordDocument.Create(ms)) {
                var paragraph = document.AddParagraph(string.Empty);
                paragraph.AddFormattedText("Hello");
                paragraph.AddFormattedText("Bold", bold: true);
                paragraph.AddFormattedText("Italic", italic: true);
                paragraph.AddFormattedText("Strike").Strike = true;
                var codeRun = paragraph.AddFormattedText("Code");
                codeRun.SetFontFamily(FontResolver.Resolve("monospace")!);
                paragraph.AddHyperLink("Link", new Uri("https://example.com/"));
                paragraph.AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));

                document.Save();

                var runs = FormattingHelper.GetFormattedRuns(paragraph).ToList();
                Assert.Equal(7, runs.Count);
                Assert.Contains(runs, r => r.Text == "Hello" && !r.Bold);
                Assert.Contains(runs, r => r.Text == "Bold" && r.Bold);
                Assert.Contains(runs, r => r.Text == "Italic" && r.Italic);
                Assert.Contains(runs, r => r.Text == "Strike" && r.Strike);
                Assert.Contains(runs, r => r.Text == "Code" && r.Code);
                Assert.Contains(runs, r => r.Text == "Link" && r.Hyperlink == "https://example.com/");
                Assert.Contains(runs, r => r.Image != null);
            }
        }

        [Fact]
        public void DocumentTraversal_ResolvesListMarkers() {
            using MemoryStream ms = new MemoryStream();
            using (var document = WordDocument.Create(ms)) {
                var bullet = document.AddList(WordListStyle.Bulleted);
                var bulletItem = bullet.AddItem("Bullet 1");
                var ordered = document.AddList(WordListStyle.Headings111);
                var orderedItem = ordered.AddItem("Number 1");

                document.Save();

                var bulletInfo = DocumentTraversal.GetListInfo(bulletItem);
                Assert.NotNull(bulletInfo);
                Assert.False(bulletInfo.Value.Ordered);

                var orderedInfo = DocumentTraversal.GetListInfo(orderedItem);
                Assert.True(orderedInfo.Value.Ordered);

                var markers = DocumentTraversal.BuildListMarkers(document);
                Assert.Equal("•", markers[bulletItem].Marker);
                Assert.Equal("1.", markers[orderedItem].Marker);
            }
        }
    }
}


using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_CreatingWordWithHyperlinks() {
            using (WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "HyperlinksTests.docx"))) {
                Assert.True(document.Paragraphs.Count == 0);
                Assert.True(document.Sections.Count == 1);
                Assert.True(document.Fields.Count == 0);
                Assert.True(document.HyperLinks.Count == 0);
                Assert.True(document.ParagraphsHyperLinks.Count == 0);
                Assert.True(document.Bookmarks.Count == 0);

                document.AddParagraph("Test 1");

                document.AddParagraph("Hello users! Please visit ").AddHyperLink("bookmark below", "TestBookmark", true, "This is link to bookmark below shown within Tooltip");

                Assert.True(document.Paragraphs.Count == 3);
                Assert.True(document.HyperLinks.Count == 1);
                Assert.True(document.ParagraphsHyperLinks.Count == 1);

                document.AddParagraph("Test HYPERLINK ").AddHyperLink(" to website?", new Uri("https://evotec.xyz"));

                Assert.True(document.Paragraphs.Count == 5);
                Assert.True(document.HyperLinks.Count == 2);
                Assert.True(document.ParagraphsHyperLinks.Count == 2);

                document.AddParagraph("Test Email Address ").AddHyperLink("Przemysław Klys", new Uri("mailto:kontakt@evotec.pl?subject=Test Subject"));

                Assert.True(document.Paragraphs.Count == 7);
                Assert.True(document.HyperLinks.Count == 3);
                Assert.True(document.ParagraphsHyperLinks.Count == 3);

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test HYPERLINK ").AddHyperLink(" to website?", new Uri("https://evotec.xyz"));
                document.AddParagraph("Test HYPERLINK ").AddHyperLink(" to website?", new Uri("https://evotec.pl"));

                Assert.True(document.Paragraphs.Count == 13);
                Assert.True(document.HyperLinks.Count == 5);
                Assert.True(document.ParagraphsHyperLinks.Count == 5);

                document.HyperLinks.Last().Remove();

                Assert.True(document.Paragraphs.Count == 12);
                Assert.True(document.HyperLinks.Count == 4);
                Assert.True(document.ParagraphsHyperLinks.Count == 4);

                document.AddParagraph("Test 2").AddBookmark("TestBookmark");

                Assert.True(document.Paragraphs.Count == 14);
                Assert.True(document.HyperLinks.Count == 4);
                Assert.True(document.ParagraphsHyperLinks.Count == 4);
                Assert.True(document.Bookmarks.Count == 1);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "HyperlinksTests.docx"))) {

            }
        }
    }
}

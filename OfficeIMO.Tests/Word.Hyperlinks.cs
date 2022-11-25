using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using SemanticComparison;
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

                var hyperlink = document.AddParagraph("Hello users! Please visit ").AddHyperLink("bookmark below", "TestBookmark", true, "This is link to bookmark below shown within Tooltip");

                Assert.True(hyperlink.Hyperlink.Text == "bookmark below");

                Assert.True(document.Paragraphs.Count == 3);
                Assert.True(document.HyperLinks.Count == 1);
                Assert.True(document.ParagraphsHyperLinks.Count == 1);

                document.AddParagraph("Test HYPERLINK ").AddHyperLink(" to website?", new Uri("https://evotec.xyz"));

                Assert.True(document.Paragraphs.Count == 5);
                Assert.True(document.HyperLinks.Count == 2);
                Assert.True(document.ParagraphsHyperLinks.Count == 2);

                var hyperlink1 = document.AddParagraph("Test Email Address ").AddHyperLink("Przemysław Klys", new Uri("mailto:kontakt@evotec.pl?subject=Test Subject"));

                Assert.True(hyperlink1.Hyperlink.EmailAddress == "kontakt@evotec.pl");


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

                Assert.True(document.Sections[0].Paragraphs.Count == 14);
                Assert.True(document.Sections[0].HyperLinks.Count == 4);
                Assert.True(document.Sections[0].ParagraphsHyperLinks.Count == 4);
                Assert.True(document.Sections[0].Bookmarks.Count == 1);

                document.Save(false);

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "HyperlinksTests.docx"))) {
                Assert.True(document.Paragraphs.Count == 14);
                Assert.True(document.HyperLinks.Count == 4);
                Assert.True(document.ParagraphsHyperLinks.Count == 4);
                Assert.True(document.Bookmarks.Count == 1);
                Assert.True(document.Sections[0].Paragraphs.Count == 14);
                Assert.True(document.Sections[0].HyperLinks.Count == 4);
                Assert.True(document.Sections[0].ParagraphsHyperLinks.Count == 4);
                Assert.True(document.Sections[0].Bookmarks.Count == 1);

                document.HyperLinks[1].Uri = new Uri("https://evotec.pl");

                Assert.True(document.HyperLinks[1].Uri == new Uri("https://evotec.pl"));

                Assert.True(document.HyperLinks[1].Text == " to website?");

                var section = document.AddSection(SectionMarkValues.NextPage);
                section.AddHyperLink("This is my website", new Uri("https://evotec.xyz"));
                section.AddHyperLink("This is second website", new Uri("https://evotec.pl"), true, "This is tooltip for my website 1");
                document.AddHyperLink("This is third website", new Uri("https://evotec.se"), true, "This is tooltip for my website 2");

                Assert.True(document.Paragraphs.Count == 17);

                Assert.True(section.HyperLinks[0].Text == "This is my website");
                Assert.True(section.HyperLinks[1].Text == "This is second website");
                Assert.True(section.HyperLinks[2].Text == "This is third website");
                Assert.True(section.HyperLinks[0].Anchor == null);
                Assert.True(section.HyperLinks[1].Anchor == null);
                Assert.True(section.HyperLinks[2].Anchor == null);
                Assert.True(section.HyperLinks[0].Uri == new Uri("https://evotec.xyz"));
                Assert.True(section.HyperLinks[1].Uri == new Uri("https://evotec.pl"));
                Assert.True(section.HyperLinks[2].Uri == new Uri("https://evotec.se"));
                Assert.True(section.HyperLinks[0].Tooltip == null);
                Assert.True(section.HyperLinks[1].Tooltip == "This is tooltip for my website 1");
                Assert.True(section.HyperLinks[2].Tooltip == "This is tooltip for my website 2");

                Assert.True(document.Sections[1].HyperLinks[0].Text == "This is my website");
                Assert.True(document.Sections[1].HyperLinks[1].Text == "This is second website");
                Assert.True(document.Sections[1].HyperLinks[2].Text == "This is third website");
                Assert.True(document.Sections[1].HyperLinks[0].Anchor == null);
                Assert.True(document.Sections[1].HyperLinks[1].Anchor == null);
                Assert.True(document.Sections[1].HyperLinks[2].Anchor == null);
                Assert.True(document.Sections[1].HyperLinks[0].Uri == new Uri("https://evotec.xyz"));
                Assert.True(document.Sections[1].HyperLinks[1].Uri == new Uri("https://evotec.pl"));
                Assert.True(document.Sections[1].HyperLinks[2].Uri == new Uri("https://evotec.se"));
                Assert.True(document.Sections[1].HyperLinks[0].Tooltip == null);
                Assert.True(document.Sections[1].HyperLinks[1].Tooltip == "This is tooltip for my website 1");
                Assert.True(document.Sections[1].HyperLinks[2].Tooltip == "This is tooltip for my website 2");

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "HyperlinksTests.docx"))) {
                Assert.True(document.Paragraphs.Count == 17);
                Assert.True(document.HyperLinks.Count == 7);
                Assert.True(document.ParagraphsHyperLinks.Count == 7);
                Assert.True(document.Bookmarks.Count == 1);
                Assert.True(document.Sections[0].Paragraphs.Count == 14);
                Assert.True(document.Sections[0].HyperLinks.Count == 4);
                Assert.True(document.Sections[0].ParagraphsHyperLinks.Count == 4);
                Assert.True(document.Sections[0].Bookmarks.Count == 1);

                var section = document.Sections[1];

                Assert.True(section.HyperLinks[0].Text == "This is my website");
                Assert.True(section.HyperLinks[1].Text == "This is second website");
                Assert.True(section.HyperLinks[2].Text == "This is third website");
                Assert.True(section.HyperLinks[0].Anchor == null);
                Assert.True(section.HyperLinks[1].Anchor == null);
                Assert.True(section.HyperLinks[2].Anchor == null);
                Assert.True(section.HyperLinks[0].Uri == new Uri("https://evotec.xyz"));
                Assert.True(section.HyperLinks[1].Uri == new Uri("https://evotec.pl"));
                Assert.True(section.HyperLinks[2].Uri == new Uri("https://evotec.se"));
                Assert.True(section.HyperLinks[0].Tooltip == null);
                Assert.True(section.HyperLinks[1].Tooltip == "This is tooltip for my website 1");
                Assert.True(section.HyperLinks[2].Tooltip == "This is tooltip for my website 2");

                Assert.True(document.Sections[1].HyperLinks[0].Text == "This is my website");
                Assert.True(document.Sections[1].HyperLinks[1].Text == "This is second website");
                Assert.True(document.Sections[1].HyperLinks[2].Text == "This is third website");
                Assert.True(document.Sections[1].HyperLinks[0].Anchor == null);
                Assert.True(document.Sections[1].HyperLinks[1].Anchor == null);
                Assert.True(document.Sections[1].HyperLinks[2].Anchor == null);
                Assert.True(document.Sections[1].HyperLinks[0].Uri == new Uri("https://evotec.xyz"));
                Assert.True(document.Sections[1].HyperLinks[1].Uri == new Uri("https://evotec.pl"));

                document.Sections[1].HyperLinks[1].History = false;

                Assert.True(document.Sections[1].HyperLinks[2].Uri == new Uri("https://evotec.se"));
                Assert.True(document.Sections[1].HyperLinks[0].Tooltip == null);
                Assert.True(document.Sections[1].HyperLinks[1].Tooltip == "This is tooltip for my website 1");
                Assert.True(document.Sections[1].HyperLinks[2].Tooltip == "This is tooltip for my website 2");

                Assert.True(document.Sections[1].HyperLinks[0].IsEmail == false);
                Assert.True(document.Sections[1].HyperLinks[0].IsHttp == true);
                Assert.True(document.Sections[1].HyperLinks[0].Scheme == Uri.UriSchemeHttps);
                Assert.True(document.Sections[1].HyperLinks[0].History == true);
                Assert.True(document.Sections[1].HyperLinks[0].TargetFrame == null);


                Assert.True(document.Sections[1].HyperLinks[1].History == false);

                document.Save();
            }
        }
    }
}

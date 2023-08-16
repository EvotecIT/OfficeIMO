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
using Color = SixLabors.ImageSharp.Color;

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

                Assert.True(hyperlink.Underline == UnderlineValues.Single);
                Assert.True(hyperlink.Bold == false);
                Assert.True(hyperlink.Italic == false);
                Assert.True(hyperlink.Color == Color.Blue);

                hyperlink.Bold = true;
                hyperlink.Italic = true;

                Assert.True(hyperlink.Bold);
                Assert.True(hyperlink.Italic);
                Assert.True(hyperlink.Underline == UnderlineValues.Single);
                Assert.True(hyperlink.Color == Color.Blue);

                hyperlink.Color = Color.Red;
                hyperlink.Underline = UnderlineValues.Dash;

                Assert.True(hyperlink.Color == Color.Red);
                Assert.True(hyperlink.Underline == UnderlineValues.Dash);

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

        [Fact]
        public void Test_CreatingWordWithHyperlinksVerification() {
            using (WordDocument document =
                   WordDocument.Create(Path.Combine(_directoryWithFiles, "HyperlinksTests.docx"))) {
                Assert.True(document.Paragraphs.Count == 0);
                Assert.True(document.Sections.Count == 1);
                Assert.True(document.Fields.Count == 0);
                Assert.True(document.HyperLinks.Count == 0);
                Assert.True(document.ParagraphsHyperLinks.Count == 0);
                Assert.True(document.Bookmarks.Count == 0);

                var paragraph = document.AddParagraph("Test 1");
                Assert.True(paragraph.Bold == false);

                var hyperlink = document.AddParagraph("Hello users! Please visit ").AddHyperLink("bookmark below",
                    "TestBookmark", true, "This is link to bookmark below shown within Tooltip");

                Assert.True(hyperlink.Hyperlink._runProperties.Bold == null);

                Assert.True(hyperlink.Underline == UnderlineValues.Single);
                Assert.True(hyperlink.Bold == false);
                Assert.True(hyperlink.Italic == false);
                Assert.True(hyperlink.Color == Color.Blue);

                hyperlink.Bold = true;
                hyperlink.Italic = true;

                Assert.True(hyperlink.Bold);
                Assert.True(hyperlink.Italic);
                Assert.True(hyperlink.Underline == UnderlineValues.Single);
                Assert.True(hyperlink.Color == Color.Blue);

                hyperlink.Color = Color.Red;
                hyperlink.Underline = UnderlineValues.Dash;


                var hyperlinkWithoutStyle = document.AddParagraph("Hello users! Please visit ").AddHyperLink("bookmark below",
                    "TestBookmark", false, "This is link to bookmark below shown within Tooltip");

                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties == null);

                hyperlinkWithoutStyle.Bold = true;
                Assert.True(hyperlinkWithoutStyle.Bold == true);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Bold != null);

                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Italic == null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Underline == null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Color == null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Spacing == null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.FontSize == null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.RunFonts == null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Highlight == null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Strike == null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.DoubleStrike == null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Caps == null);

                hyperlinkWithoutStyle.Bold = true;
                hyperlinkWithoutStyle.Italic = true;
                hyperlinkWithoutStyle.Color = Color.Red;
                hyperlinkWithoutStyle.Underline = UnderlineValues.Dash;
                hyperlinkWithoutStyle.Spacing = 2;
                hyperlinkWithoutStyle.FontSize = 12;
                hyperlinkWithoutStyle.FontFamily = "Arial";
                hyperlinkWithoutStyle.Highlight = HighlightColorValues.Cyan;
                hyperlinkWithoutStyle.Strike = true;
                hyperlinkWithoutStyle.DoubleStrike = true;
                hyperlinkWithoutStyle.CapsStyle = CapsStyle.SmallCaps;

                Assert.True(hyperlinkWithoutStyle.Bold);
                Assert.True(hyperlinkWithoutStyle.Italic);
                Assert.True(hyperlinkWithoutStyle.Underline == UnderlineValues.Dash);
                Assert.True(hyperlinkWithoutStyle.Color == Color.Red);
                Assert.True(hyperlinkWithoutStyle.Spacing == 2);
                Assert.True(hyperlinkWithoutStyle.FontSize == 12);
                Assert.True(hyperlinkWithoutStyle.FontFamily == "Arial");
                Assert.True(hyperlinkWithoutStyle.Highlight == HighlightColorValues.Cyan);
                Assert.True(hyperlinkWithoutStyle.Strike);
                Assert.True(hyperlinkWithoutStyle.DoubleStrike);
                Assert.True(hyperlinkWithoutStyle.CapsStyle == CapsStyle.SmallCaps);

                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Bold != null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Italic != null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Underline != null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Color != null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Spacing != null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.FontSize != null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.RunFonts != null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Highlight != null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Strike != null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.DoubleStrike != null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Caps == null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.SmallCaps != null);

                hyperlinkWithoutStyle.CapsStyle = CapsStyle.Caps;

                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.Caps != null);
                Assert.True(hyperlinkWithoutStyle.Hyperlink._runProperties.SmallCaps == null);

                Assert.True(hyperlinkWithoutStyle.CapsStyle == CapsStyle.Caps);

            }
        }

        [Fact]
        public void Test_CreatingWordWithHyperlinksInTables() {
            using (WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "HyperlinksTestsInTables.docx"))) {
                Assert.True(document.Paragraphs.Count == 0);
                Assert.True(document.Sections.Count == 1);
                Assert.True(document.Fields.Count == 0);
                Assert.True(document.HyperLinks.Count == 0);
                Assert.True(document.ParagraphsHyperLinks.Count == 0);
                Assert.True(document.Bookmarks.Count == 0);

                document.AddParagraph("Test 1");

                document.AddTable(3, 3);

                document.Tables[0].Rows[0].Cells[0].Paragraphs[0].AddHyperLink(" to website?", new Uri("https://evotec.xyz"), addStyle: true);

                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs[1].IsHyperLink == true);
                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs[1].Hyperlink.IsHttp == true);
                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs[1].Hyperlink.Text == " to website?");

                Assert.True(document.HyperLinks.Count == 1);

                Assert.True(document.HyperLinks[0].Text == " to website?");

                document.Save(false);
            }
        }

    }
}

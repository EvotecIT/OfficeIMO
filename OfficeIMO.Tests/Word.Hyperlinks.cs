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

        [Fact]
        public void Test_RemoveHyperLinkMethod() {
            string filePath = Path.Combine(_directoryWithFiles, "HyperlinkRemove.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Text before ");
                paragraph.AddHyperLink("link", new Uri("https://evotec.xyz"));
                paragraph.AddText(" after");

                Assert.Single(document.HyperLinks);
                Assert.Single(document._wordprocessingDocument.MainDocumentPart.HyperlinkRelationships);

                paragraph.RemoveHyperLink();

                Assert.Empty(document.HyperLinks);
                Assert.Empty(document._wordprocessingDocument.MainDocumentPart.HyperlinkRelationships);
                Assert.Equal(2, paragraph._paragraph.ChildElements.OfType<Run>().Count());

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.HyperLinks);
                Assert.Equal(2, document.Paragraphs[0]._paragraph.ChildElements.OfType<Run>().Count());
                document.Save();
            }
        }

        [Fact]
        public void Test_CreateFormattedHyperlink() {
            string filePath = Path.Combine(_directoryWithFiles, "FormattedHyperlink.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Link to ");
                paragraph.AddHyperLink("Google", new Uri("https://google.com"), addStyle: true);

                var reference = paragraph.Hyperlink;
                var created = WordHyperLink.CreateFormattedHyperlink(reference, "Bing", new Uri("https://bing.com"));

                Assert.Equal("Bing", created.Text);
                Assert.Equal(new Uri("https://bing.com"), created.Uri);
                Assert.NotNull(created._runProperties);
                Assert.NotNull(created._runProperties.Color);
                Assert.Equal(reference._runProperties.Color.Val, created._runProperties.Color.Val);
                Assert.Equal(reference._runProperties.Color.ThemeColor, created._runProperties.Color.ThemeColor);
                Assert.Equal(reference._runProperties.Underline.Val, created._runProperties.Underline.Val);
                Assert.Equal(reference._runProperties.RunStyle.Val, created._runProperties.RunStyle.Val);
                Assert.Equal(2, paragraph._paragraph.Elements<Hyperlink>().Count());

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Paragraphs[0]._paragraph.Elements<Hyperlink>().Count());
                var secondLink = document.Paragraphs[0]._paragraph.Elements<Hyperlink>().ElementAt(1);
                Assert.Equal("Bing", secondLink.InnerText);
                document.Save();
            }
        }

        [Fact]
        public void Test_DuplicateHyperlink() {
            string filePath = Path.Combine(_directoryWithFiles, "DuplicateHyperlink.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Search using ");
                paragraph.AddHyperLink("Google", new Uri("https://google.com"), addStyle: true);
                var reference = paragraph.Hyperlink;

                var duplicate = WordHyperLink.DuplicateHyperlink(reference);

                Assert.Equal(reference.Text, duplicate.Text);
                Assert.Equal(reference.Uri, duplicate.Uri);
                Assert.Equal(2, paragraph._paragraph.Elements<Hyperlink>().Count());

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Paragraphs[0]._paragraph.Elements<Hyperlink>().Count());
                document.Save();
            }
        }

        [Fact]
        public void Test_InsertFormattedHyperlinkBefore() {
            string filePath = Path.Combine(_directoryWithFiles, "FormattedBefore.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Link to ");
                paragraph.AddHyperLink("Google", new Uri("https://google.com"), addStyle: true);

                var reference = paragraph.Hyperlink;
                var created = reference.InsertFormattedHyperlinkBefore("Bing", new Uri("https://bing.com"));

                Assert.Equal(2, paragraph._paragraph.Elements<Hyperlink>().Count());
                Assert.Equal("Bing", paragraph._paragraph.Elements<Hyperlink>().First().InnerText);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                var firstLink = document.Paragraphs[0]._paragraph.Elements<Hyperlink>().First();
                Assert.Equal("Bing", firstLink.InnerText);
                document.Save();
            }
        }

        [Fact]
        public void Test_InsertFormattedHyperlinkInHeaderAndFooter() {
            string filePath = Path.Combine(_directoryWithFiles, "FormattedHeaderFooter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var header = document.Header.Default;
                var paraHeader = header.AddParagraph("Search using ");
                paraHeader.AddHyperLink("Google", new Uri("https://google.com"), addStyle: true);
                var refHeader = paraHeader.Hyperlink;
                refHeader.InsertFormattedHyperlinkAfter("Bing", new Uri("https://bing.com"));

                var footer = document.Footer.Default;
                var paraFooter = footer.AddParagraph("Find us on ");
                paraFooter.AddHyperLink("Yahoo", new Uri("https://yahoo.com"), addStyle: true);
                var refFooter = paraFooter.Hyperlink;
                refFooter.InsertFormattedHyperlinkAfter("DuckDuckGo", new Uri("https://duckduckgo.com"));

                Assert.Equal(2, paraHeader._paragraph.Elements<Hyperlink>().Count());
                Assert.Equal(2, paraFooter._paragraph.Elements<Hyperlink>().Count());

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                var headerPara = document.Header.Default.Paragraphs[0];
                Assert.Equal(2, headerPara._paragraph.Elements<Hyperlink>().Count());
                var footerPara = document.Footer.Default.Paragraphs[0];
                Assert.Equal(2, footerPara._paragraph.Elements<Hyperlink>().Count());
                document.Save();
            }
        }

        [Fact]
        public void Test_CopyFormattingFromHyperlink() {
            string filePath = Path.Combine(_directoryWithFiles, "CopyFormatting.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Go to ");
                var refPara = paragraph.AddHyperLink("Google", new Uri("https://google.com"), addStyle: true);
                refPara.Bold = true;
                var reference = refPara.Hyperlink;

                paragraph.AddHyperLink("Bing", new Uri("https://bing.com"));
                var target = paragraph.Hyperlink;
                target.CopyFormattingFrom(reference);

                Assert.NotNull(target._runProperties);
                Assert.Equal(reference._runProperties.Color.Val, target._runProperties.Color.Val);
                Assert.Equal(reference._runProperties.Underline.Val, target._runProperties.Underline.Val);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Paragraphs[0]._paragraph.Elements<Hyperlink>().Count());
                document.Save();
            }
        }

        [Fact]
        public void Test_ListHyperlinkFormattingReuse() {
            string filePath = Path.Combine(_directoryWithFiles, "ListFormatting.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var first = document.AddList(WordListStyle.Bulleted);
                var google = first.AddItem("").AddHyperLink("Google", new Uri("https://google.com"), addStyle: true);
                google.Bold = true;
                var googleRef = google.Hyperlink;

                var bing = first.AddItem("").AddHyperLink("Bing", new Uri("https://bing.com"), addStyle: true);
                bing.Italic = true;
                var bingRef = bing.Hyperlink;

                document.AddParagraph("separator");

                var second = document.AddList(WordListStyle.Bulleted);
                var duck = second.AddItem("").AddHyperLink("DuckDuckGo", new Uri("https://duckduckgo.com"));
                duck.Hyperlink.CopyFormattingFrom(googleRef);
                var start = second.AddItem("").AddHyperLink("Startpage", new Uri("https://startpage.com"));
                start.Hyperlink.CopyFormattingFrom(bingRef);

                Assert.True(duck.Bold);
                Assert.True(start.Italic);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Lists.Count);
                document.Save();
            }
        }

    }
}

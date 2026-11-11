using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;
using Xunit;
using Path = System.IO.Path;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_Range_ReplaceAcrossWordAuthoredRunsKeepsFirstCharacterFormatting() {
            // Word wrote "This is 2" / superscript "nd" / " page." as separate runs.
            string filePath = CopyFixtureDoc("FieldsAndSections.docx");
            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRange range = Assert.Single(document.Find("nd page").Ranges);
                Assert.Equal("nd page", range.Text);
                range.Replace("nd edition");
                Assert.Equal("nd edition", range.Text);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());
                Paragraph paragraph = document._document.Body!.Elements<Paragraph>()
                    .Single(p => p.InnerText.Contains("edition"));
                Assert.Equal("This is 2nd edition.", paragraph.InnerText);
                var runs = paragraph.Elements<Run>().Where(run => run.InnerText.Length > 0).ToList();
                Assert.Equal(new[] { "This is 2", "nd edition", "." }, runs.Select(run => run.InnerText));
                Assert.Null(runs[0].RunProperties?.VerticalTextAlignment);
                Assert.Equal(VerticalPositionValues.Superscript, runs[1].RunProperties!.VerticalTextAlignment!.Val!.Value);
                Assert.Null(runs[2].RunProperties?.VerticalTextAlignment);
            }
        }

        [Fact]
        public void Test_Range_RegexReplaceAllSubstitutesGroupsInEveryStory() {
            string filePath = Path.Combine(_directoryWithFiles, "RangeRegexReplace.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph body = document.AddParagraph("Order 1");
                body.AddText("2-3").SetBold();
                body.AddText("4 shipped");
                body.AddFootNote("Note 11-22");
                WordTable table = document.AddTable(1, 1);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Cell 56-78";
                document.PageSetup.Header.AddParagraph("Header 90-12");

                var pattern = new Regex(@"(\d+)-(\d+)");
                Assert.Equal(4, document.Find(pattern).MatchCount);
                Assert.Equal(4, document.ReplaceAll(pattern, "$2/$1"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());
                Assert.Equal("Order 34/12 shipped", document._document.Body!.Elements<Paragraph>().First().InnerText);
                Assert.Equal("Cell 78/56", document.Tables[0].Rows[0].Cells[0].Paragraphs[0]._paragraph.InnerText);
                Assert.Contains("Header 12/90", document.PageSetup.GetHeader(WordHeaderFooterType.Default)!.Paragraphs.Select(p => p._paragraph.InnerText));
                Assert.Contains("Note 22/11", document.FootNotes.Select(note => string.Concat(note.Paragraphs!.Select(p => p._paragraph.InnerText).Distinct())));
            }
        }

        [Fact]
        public void Test_Range_RegexCallbacksReplaceWithTextHyperlinksAndImages() {
            string filePath = Path.Combine(_directoryWithFiles, "RangeRegexCallbacks.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("status: open, owner: kim");
                document.AddParagraph("See [the guide](https://evotec.xyz/guide) and [news](https://evotec.xyz/news).");
                document.AddParagraph("Logo: {{logo}}");

                Assert.Equal(2, document.ReplaceAll(new Regex(@"(?<=: )\w+"), match => match.Value.ToUpperInvariant()));
                Assert.Equal(2, document.ReplaceAll(new Regex(@"\[(.+?)\]\((.+?)\)"),
                    (match, paragraph) => paragraph.AddHyperLink(match.Groups[1].Value, new Uri(match.Groups[2].Value), addStyle: true)));
                document.Find("{{logo}}").Ranges.Single()
                    .Replace(paragraph => paragraph.AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), OfficeLength.Points(40), OfficeLength.Points(20)));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());
                var paragraphs = document._document.Body!.Elements<Paragraph>().ToList();
                Assert.Equal("status: OPEN, owner: KIM", paragraphs[0].InnerText);
                Assert.Equal("See the guide and news.", paragraphs[1].InnerText);
                Assert.Equal(new[] { "https://evotec.xyz/guide", "https://evotec.xyz/news" },
                    document.HyperLinks.Select(link => link.Uri!.ToString()));
                Assert.Equal("Logo: ", paragraphs[2].InnerText);
                Assert.Single(document.Images);
            }
        }

        [Fact]
        public void Test_Range_BookmarkSpansTextAndRefFieldsFollowItsText() {
            string filePath = Path.Combine(_directoryWithFiles, "RangeBookmarks.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph value = document.AddParagraph("The contract value is ");
                value.AddText("1,000").SetBold();
                value.AddText(" EUR today.");
                document.Find("1,000 EUR").Ranges.Single().AddBookmark("Amount");
                document.AddParagraph("Total: ").AddCrossReference("Amount", WordCrossReferenceType.Bookmark);
                WordBookmark empty = document.AddParagraph("Signed by ").AddBookmark("Signer");

                Assert.Equal("1,000 EUR", document.Bookmarks.Single(b => b.Name == "Amount").Text);
                Assert.Contains(document.UpdateFieldsAndGetReport().Results, result => result.ResultText == "1,000 EUR");

                document.Bookmarks.Single(b => b.Name == "Amount").Text = "2,500 EUR";
                empty.Text = "Przemysław";
                Assert.Contains(document.UpdateFieldsAndGetReport().Results, result => result.ResultText == "2,500 EUR");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());
                var paragraphs = document._document.Body!.Elements<Paragraph>().ToList();
                Assert.Equal("The contract value is 2,500 EUR today.", paragraphs[0].InnerText);
                Assert.Equal("2,500 EUR", document.Bookmarks.Single(b => b.Name == "Amount").Text);
                Assert.Equal("Przemysław", document.Bookmarks.Single(b => b.Name == "Signer").Text);
                Assert.Equal("Signed by Przemysław", paragraphs[2].InnerText);
            }
        }

        [Fact]
        public void Test_Range_CommentAcrossParagraphsWithRichBody() {
            string filePath = Path.Combine(_directoryWithFiles, "RangeComments.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph first = document.AddParagraph("Alpha beta");
                WordParagraph second = document.AddParagraph("gamma delta");
                WordRange range = first.GetRange(6, 4).ExtendTo(second.GetRange(0, 5));
                Assert.Equal("beta\ngamma", range.Text);
                range.AddComment("Przemysław Kłys", "PK", body => {
                    body.AddText("Check ");
                    body.AddText("this").SetBold();
                    body.AddHyperLink(" spec", new Uri("https://evotec.xyz/spec"));
                });
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());
                WordComment comment = Assert.Single(document.Comments);
                Assert.Equal("Check this spec", comment.Text);
                var paragraphs = document._document.Body!.Elements<Paragraph>().ToList();
                Assert.Equal("Alpha ", string.Concat(paragraphs[0].Elements().TakeWhile(e => e is not CommentRangeStart).Select(e => e.InnerText)));
                Assert.Equal("gamma", string.Concat(paragraphs[1].Elements().TakeWhile(e => e is not CommentRangeEnd).Select(e => e.InnerText)));
                Assert.Single(paragraphs[1].Descendants<CommentReference>());

                var commentsPart = document._wordprocessingDocument.MainDocumentPart!.WordprocessingCommentsPart!;
                Comment stored = commentsPart.Comments!.Elements<Comment>().Single();
                Assert.Contains(stored.Descendants<Run>(), run => run.InnerText == "this" && run.RunProperties?.Bold != null);
                Hyperlink link = stored.Descendants<Hyperlink>().Single();
                Assert.Equal("https://evotec.xyz/spec", commentsPart.HyperlinkRelationships.Single(r => r.Id == link.Id).Uri.ToString());
            }
        }

        [Fact]
        public void Test_Range_HyperlinkWrapsOnlyTheRangeAndKeepsRunFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "RangeHyperlink.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph paragraph = document.AddParagraph("Visit ");
                paragraph.AddText("Office").SetBold();
                paragraph.AddText("IMO docs now");
                document.Find("OfficeIMO docs").Ranges.Single().AddHyperLink(new Uri("https://officeimo.com"));
                // Replacing all of an existing link's text replaces the link instead of nesting a new one inside it.
                document.AddParagraph("Old: ").AddHyperLink("old site", new Uri("https://example.com/old"));
                document.Find("old site").Ranges.Single().Replace(p => p.AddHyperLink("new site", new Uri("https://example.com/new")));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());
                Paragraph paragraph = document._document.Body!.Elements<Paragraph>().First();
                Assert.Equal("Visit OfficeIMO docs now", paragraph.InnerText);
                Hyperlink link = paragraph.Elements<Hyperlink>().Single();
                Assert.Equal("OfficeIMO docs", link.InnerText);
                Assert.NotNull(link.Elements<Run>().First().RunProperties!.Bold);
                Assert.Equal(new[] { "https://officeimo.com/", "https://example.com/new" }, document.HyperLinks.Select(h => h.Uri!.ToString()));
                Paragraph replaced = document._document.Body!.Elements<Paragraph>().ElementAt(1);
                Assert.Equal("Old: new site", replaced.InnerText);
                Assert.Empty(replaced.Descendants<Hyperlink>().Single().Descendants<Hyperlink>());
            }
        }

        [Fact]
        public void Test_Range_ReplacingRangesInOrderAndAcrossParagraphs() {
            string filePath = Path.Combine(_directoryWithFiles, "RangeSequence.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("a-x-x");
                document.AddParagraph("x-b");
                document.AddParagraph("Hello wor");
                document.AddParagraph("ld again");

                foreach (WordRange range in document.Find("x").Ranges) range.Replace("long");
                WordRange world = document.Find("world").Ranges.Single();
                Assert.Equal("wor\nld", world.Text);
                world.Replace("planet");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());
                Assert.Equal(new[] { "a-long-long", "long-b", "Hello planet again" },
                    document._document.Body!.Elements<Paragraph>().Select(p => p.InnerText));
            }
        }

        [Fact]
        public void Test_Range_FieldResultSpanningParagraphsIsNotEditable() {
            using WordDocument document = WordDocument.Create();
            Body body = document._document.Body!;
            document.AddParagraph("Contents ")._paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" TOC \\o \"1-3\" ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("Entry one")));
            document.AddParagraph("Entry two");
            document.AddParagraph("Entry three")._paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new Text(" after") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
            string before = body.InnerText;

            // The middle paragraph has no field markers of its own; its text is still the field result.
            Assert.Throws<InvalidOperationException>(() => document.Find("Entry two").Ranges.Single().Replace("x"));
            Paragraph middle = body.Elements<Paragraph>().ElementAt(1);
            Assert.Throws<InvalidOperationException>(() => new WordParagraph(document, middle).GetRange().Replace("x"));
            Assert.Equal(before, body.InnerText);

            // Text after the field end stays editable.
            document.Find(" after").Ranges.Single().Replace(" later");
            Assert.EndsWith("Entry three later", body.Elements<Paragraph>().Last().InnerText, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_Range_UnsafeEditsThrowAndLeaveTextUnchanged() {
            string filePath = CopyFixtureDoc("FieldsAndSections.docx");
            using (WordDocument document = WordDocument.Load(filePath)) {
                string before = document._document.Body!.InnerText;

                // Complex field result, simple field result, and a match list that includes a field result.
                Assert.Throws<InvalidOperationException>(() => document.Find("FieldsAndSections.docx").Ranges.Single().Replace("x"));
                Assert.Throws<InvalidOperationException>(() => document.Find("Przemysław Kłys").Ranges.Single().Replace("x"));
                Assert.Throws<InvalidOperationException>(() => document.ReplaceAll(new Regex("other|Sections"), "X"));
                // An empty range right after a field result would insert into the field.
                WordParagraph fieldParagraph = document.Paragraphs.First(p => p.Text.Contains("Lets"));
                int afterField = fieldParagraph.GetRange().Text.IndexOf(" and continue", StringComparison.Ordinal);
                Assert.Throws<InvalidOperationException>(() => fieldParagraph.GetRange(afterField, 0).Replace("x"));
                // Joining paragraphs across a section break.
                WordRange page = document.Find("page.").Ranges.Last();
                WordRange section = document.Find("This is 1").Ranges.Single();
                Assert.Throws<InvalidOperationException>(() => page.ExtendTo(section).Replace("x"));
                // Word does not allow comments in headers.
                WordParagraph header = document.Sections[0].PageSetup.GetHeader(WordHeaderFooterType.Default)!.Paragraphs.First(p => p.Text.Length > 0);
                Assert.Throws<InvalidOperationException>(() => header.GetRange().AddComment("A", "A", "x"));

                Assert.Equal(before, document._document.Body!.InnerText);
            }

            using (WordDocument document = WordDocument.Create()) {
                Paragraph tracked = document.AddParagraph("Kept ")._paragraph;
                tracked.Append(new InsertedRun(new Run(new Text("inserted"))) { Id = "1", Author = "A" });
                WordParagraph control = document.AddParagraph("Before ");
                control.AddStructuredDocumentTag("inside", tag: "cc");

                Assert.Throws<InvalidOperationException>(() => document.Find("Kept inserted").Ranges.Single().Replace("x"));
                Assert.Throws<InvalidOperationException>(() => document.Find("Before inside").Ranges.Single().Replace("x"));
                document.Find("inside").Ranges.Single().Replace("changed");
                Assert.Equal("changed", document.GetStructuredDocumentTagByTag("cc")!.Text);
            }
        }
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_SplitBySectionKeepsInheritedHeadersAndOnlyUsedParts() {
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            using var source = WordDocument.Create();
            _ = source.PageSetup.Header;
            source.PageSetup.GetHeader(WordHeaderFooterType.Default)!.AddParagraph("Report header");
            source.AddParagraph("Cover").AddImage(imagePath, OfficeLength.Pixels(32), OfficeLength.Pixels(32));
            source.AddParagraph("Cover note").AddFootNote("Cover footnote");

            WordSection landscape = source.AddSection();
            landscape.PageSetup.Orientation = OfficePageOrientation.Landscape;
            source.AddParagraph("Findings").AddComment("Reviewer", "R", "Check findings");
            WordList list = source.AddList(WordListStyle.Numbered);
            list.AddItem("First finding");

            source.AddSection();
            source.AddParagraph("Appendix").AddEndNote("Appendix endnote");

            IReadOnlyList<WordDocument> pieces = source.SplitBySection();
            try {
                Assert.Equal(3, pieces.Count);

                MainDocumentPart cover = Reload(pieces[0], out WordDocument coverDocument);
                MainDocumentPart findings = Reload(pieces[1], out WordDocument findingsDocument);
                MainDocumentPart appendix = Reload(pieces[2], out WordDocument appendixDocument);
                using (coverDocument)
                using (findingsDocument)
                using (appendixDocument) {
                    Assert.Equal(new[] { "Cover", "Cover note" }, Texts(coverDocument));
                    Assert.Equal(new[] { "Findings", "First finding" }, Texts(findingsDocument));
                    Assert.Equal(new[] { "Appendix" }, Texts(appendixDocument));

                    // Sections that inherit the report header in the source still show it.
                    foreach (WordDocument piece in new[] { coverDocument, findingsDocument, appendixDocument }) {
                        Assert.Single(piece.Sections);
                        Assert.Contains("Report header", piece.Sections[0].PageSetup.GetHeader(WordHeaderFooterType.Default)!.Paragraphs.Select(p => p.Text));
                    }
                    Assert.Equal(OfficePageOrientation.Landscape, findingsDocument.Sections[0].PageSetup.Size.Orientation);

                    // Each piece carries only the parts, notes, comments and lists its own content uses.
                    Assert.Single(cover.ImageParts);
                    Assert.Empty(findings.ImageParts);
                    Assert.Empty(appendix.ImageParts);
                    Assert.Single(cover.HeaderParts);
                    Assert.Single(UserNotes(cover.FootnotesPart?.Footnotes));
                    Assert.Empty(UserNotes(appendix.FootnotesPart?.Footnotes));
                    Assert.Single(UserNotes(appendix.EndnotesPart?.Endnotes));
                    Assert.Empty(UserNotes(cover.EndnotesPart?.Endnotes));
                    Assert.Equal("Check findings", Assert.Single(findings.WordprocessingCommentsPart!.Comments!.Elements<Comment>()).InnerText);
                    Assert.Empty(cover.WordprocessingCommentsPart?.Comments?.Elements<Comment>() ?? Enumerable.Empty<Comment>());
                    Assert.Single(findings.NumberingDefinitionsPart!.Numbering!.Elements<NumberingInstance>());
                    Assert.Empty(cover.NumberingDefinitionsPart?.Numbering?.Elements<NumberingInstance>() ?? Enumerable.Empty<NumberingInstance>());

                    AssertMergedDocumentIsValid(coverDocument);
                    AssertMergedDocumentIsValid(findingsDocument);
                    AssertMergedDocumentIsValid(appendixDocument);
                }
            } finally {
                foreach (WordDocument piece in pieces) piece.Dispose();
            }
        }

        [Fact]
        public void Test_SplitByHeadingStartsAtLevelAndTrimsMarkersCrossingPieces() {
            using var source = WordDocument.Create();
            source.AddParagraph("Report title");
            source.AddParagraph("Chapter A").Style = WordParagraphStyles.Heading1;
            WordParagraph opening = source.AddParagraph("Opening text");
            source.AddParagraph("Detail A.1").Style = WordParagraphStyles.Heading2;
            source.AddParagraph("Chapter B").Style = WordParagraphStyles.Heading1;
            WordParagraph closing = source.AddParagraph("Closing text");

            // A bookmark and a complex field that start in chapter A and end in chapter B.
            Paragraph openingParagraph = Body(source).Elements<Paragraph>().Single(paragraph => paragraph.InnerText == opening.Text);
            Paragraph closingParagraph = Body(source).Elements<Paragraph>().Single(paragraph => paragraph.InnerText == closing.Text);
            openingParagraph.InsertAfter(new BookmarkStart { Id = "7", Name = "Span" }, openingParagraph.ParagraphProperties);
            openingParagraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" QUOTE \"spanning\" ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
            closingParagraph.InsertAfter(new Run(new FieldChar { FieldCharType = FieldCharValues.End }), closingParagraph.ParagraphProperties);
            closingParagraph.Append(new BookmarkEnd { Id = "7" });

            IReadOnlyList<WordDocument> chapters = source.SplitByHeading();
            IReadOnlyList<WordDocument> sections = source.SplitByHeading(2);
            try {
                Assert.Equal(new[] { "Report title" }, Texts(chapters[0]));
                Assert.Equal(new[] { "Chapter A", "Opening text", "Detail A.1" }, Texts(chapters[1]));
                Assert.Equal(new[] { "Chapter B", "Closing text" }, Texts(chapters[2]));
                Assert.Equal(3, chapters.Count);
                Assert.Equal(new[] { "Detail A.1" }, Texts(sections[2]));
                Assert.Equal(4, sections.Count);

                foreach (WordDocument piece in chapters.Concat(sections)) {
                    Body body = Body(piece);
                    Assert.Empty(body.Descendants<BookmarkStart>());
                    Assert.Empty(body.Descendants<BookmarkEnd>());
                    var marks = body.Descendants<FieldChar>().Select(mark => mark.FieldCharType!.Value).ToList();
                    Assert.Equal(marks.Count(mark => mark == FieldCharValues.Begin), marks.Count(mark => mark == FieldCharValues.End));
                    AssertMergedDocumentIsValid(piece);
                }

                Styles chapterStyles = chapters[2]._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                Assert.Contains(chapterStyles.Elements<Style>(), style => style.StyleId == "Heading1");
                Assert.DoesNotContain(chapterStyles.Elements<Style>(), style => style.StyleId == "Heading2");
            } finally {
                foreach (WordDocument piece in chapters.Concat(sections)) piece.Dispose();
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_SplitReadsProtectedSourceAndKeepsImportedPreamble(bool encrypted) {
            const string password = "split-fixture";
            byte[] package;
            using (WordDocument document = WordDocument.Create()) {
                MainDocumentPart main = document._wordprocessingDocument.MainDocumentPart!;
                AlternativeFormatImportPart chunk = main.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
                using (var html = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("<html><body><p>Imported</p></body></html>"))) chunk.FeedData(html);
                Body(document).PrependChild(new AltChunk { Id = main.GetIdOfPart(chunk) });
                document.AddParagraph("Chapter").Style = WordParagraphStyles.Heading1;
                package = document.ToBytes(encrypted ? new WordSaveOptions { Password = password } : null);
            }

            using var stream = new MemoryStream(package, writable: false);
            using WordDocument source = WordDocument.Load(stream, new WordLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly,
                Password = encrypted ? password : null
            });
            IReadOnlyList<WordDocument> sections = source.SplitBySection();
            IReadOnlyList<WordDocument> pieces = source.SplitByHeading();
            try {
                Assert.Equal(new[] { "Chapter" }, Texts(Assert.Single(sections)));
                Assert.Single(Body(sections[0]).Elements<AltChunk>());
                Assert.Equal(2, pieces.Count);
                MainDocumentPart preamble = pieces[0]._wordprocessingDocument.MainDocumentPart!;
                AltChunk kept = Assert.Single(Body(pieces[0]).Elements<AltChunk>());
                Assert.IsType<AlternativeFormatImportPart>(preamble.GetPartById(kept.Id!.Value!));
                Assert.Empty(pieces[1]._wordprocessingDocument.MainDocumentPart!.AlternativeFormatImportParts);
                // The pieces are plain packages that save and reopen without the source password.
                using var reopened = WordDocument.Load(new MemoryStream(pieces[1].ToBytes(), writable: false));
                Assert.Equal(new[] { "Chapter" }, Texts(reopened));
            } finally {
                foreach (WordDocument piece in sections.Concat(pieces)) piece.Dispose();
            }
        }

        private static MainDocumentPart Reload(WordDocument piece, out WordDocument reloaded) {
            using var stream = new MemoryStream(piece.ToBytes(), writable: false);
            reloaded = WordDocument.Load(stream);
            return reloaded._wordprocessingDocument.MainDocumentPart!;
        }

        private static Body Body(WordDocument document) => document._wordprocessingDocument.MainDocumentPart!.Document!.Body!;

        private static string[] Texts(WordDocument document) =>
            document.Paragraphs.Select(paragraph => paragraph.Text).Where(text => !string.IsNullOrEmpty(text)).ToArray();

        private static IEnumerable<FootnoteEndnoteType> UserNotes(OpenXmlElement? notes) =>
            notes?.Elements<FootnoteEndnoteType>().Where(note => note.Type == null) ?? Enumerable.Empty<FootnoteEndnoteType>();
    }
}

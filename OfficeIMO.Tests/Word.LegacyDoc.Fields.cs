using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocSimpleFieldBookmarksAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph bodyParagraph = document.AddParagraph("Body ");
                    AppendBookmarkedSimpleField(bodyParagraph._paragraph, "81", "BodySimpleFieldBookmark", " PAGE ");

                    WordTable table = document.AddTable(1, 1);
                    WordParagraph cellParagraph = table.Rows[0].Cells[0].AddParagraph("Cell ", removeExistingParagraphs: true);
                    AppendBookmarkedSimpleField(cellParagraph._paragraph, "82", "CellSimpleFieldBookmark", " NUMPAGES ");

                    document.AddHeadersAndFooters();
                    WordSection section = document.Sections[0];
                    WordParagraph headerParagraph = section.Header.Default!.AddParagraph("Header ");
                    AppendBookmarkedSimpleField(headerParagraph._paragraph, "83", "HeaderSimpleFieldBookmark", " PAGE ");

                    WordParagraph footerParagraph = section.Footer.Default!.AddParagraph("Footer ");
                    AppendBookmarkedSimpleField(footerParagraph._paragraph, "84", "FooterSimpleFieldBookmark", " NUMPAGES ");

                    WordParagraph noteReferences = document.AddParagraph("Notes ");
                    WordParagraph footnoteReference = noteReferences.AddFootNote("footnote placeholder");
                    WordParagraph footnoteBody = footnoteReference.FootNote!.Paragraphs![1];
                    AppendBookmarkedSimpleField(footnoteBody._paragraph, "85", "FootnoteSimpleFieldBookmark", " PAGE ");

                    WordParagraph endnoteReference = noteReferences.AddEndNote("endnote placeholder");
                    WordParagraph endnoteBody = endnoteReference.EndNote!.Paragraphs![1];
                    AppendBookmarkedSimpleField(endnoteBody._paragraph, "86", "EndnoteSimpleFieldBookmark", " NUMPAGES ");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Empty(reloaded.LegacyDocUnsupportedFeatures);
                AssertBookmarkRoundTrip(reloaded, "BodySimpleFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "CellSimpleFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "HeaderSimpleFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "FooterSimpleFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "FootnoteSimpleFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "EndnoteSimpleFieldBookmark");
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocComplexFieldBookmarksAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph bodyParagraph = document.AddParagraph("Body ");
                    AppendBookmarkedComplexField(bodyParagraph._paragraph, "91", "BodyComplexFieldBookmark", " PAGE ");

                    WordTable table = document.AddTable(1, 1);
                    WordParagraph cellParagraph = table.Rows[0].Cells[0].AddParagraph("Cell ", removeExistingParagraphs: true);
                    AppendBookmarkedComplexField(cellParagraph._paragraph, "92", "CellComplexFieldBookmark", " NUMPAGES ");

                    document.AddHeadersAndFooters();
                    WordSection section = document.Sections[0];
                    WordParagraph headerParagraph = section.Header.Default!.AddParagraph("Header ");
                    AppendBookmarkedComplexField(headerParagraph._paragraph, "93", "HeaderComplexFieldBookmark", " PAGE ");

                    WordParagraph footerParagraph = section.Footer.Default!.AddParagraph("Footer ");
                    AppendBookmarkedComplexField(footerParagraph._paragraph, "94", "FooterComplexFieldBookmark", " NUMPAGES ");

                    WordParagraph noteReferences = document.AddParagraph("Notes ");
                    WordParagraph footnoteReference = noteReferences.AddFootNote("footnote placeholder");
                    WordParagraph footnoteBody = footnoteReference.FootNote!.Paragraphs![1];
                    AppendBookmarkedComplexField(footnoteBody._paragraph, "95", "FootnoteComplexFieldBookmark", " PAGE ");

                    WordParagraph endnoteReference = noteReferences.AddEndNote("endnote placeholder");
                    WordParagraph endnoteBody = endnoteReference.EndNote!.Paragraphs![1];
                    AppendBookmarkedComplexField(endnoteBody._paragraph, "96", "EndnoteComplexFieldBookmark", " NUMPAGES ");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Empty(reloaded.LegacyDocUnsupportedFeatures);
                AssertBookmarkRoundTrip(reloaded, "BodyComplexFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "CellComplexFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "HeaderComplexFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "FooterComplexFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "FootnoteComplexFieldBookmark");
                AssertBookmarkRoundTrip(reloaded, "EndnoteComplexFieldBookmark");
            } finally {
                DeleteIfExists(docPath);
            }
        }

        private static void AppendBookmarkedSimpleField(Paragraph paragraph, string id, string name, string instruction) {
            var simpleField = new SimpleField { Instruction = instruction };
            simpleField.Append(
                new BookmarkStart { Id = id, Name = name },
                new Run(new Text("1") { Space = SpaceProcessingModeValues.Preserve }),
                new BookmarkEnd { Id = id });
            paragraph.Append(simpleField);
        }

        private static void AppendBookmarkedComplexField(Paragraph paragraph, string id, string name, string instruction) {
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new BookmarkStart { Id = id, Name = name },
                new Run(new Text("1") { Space = SpaceProcessingModeValues.Preserve }),
                new BookmarkEnd { Id = id },
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }
    }
}

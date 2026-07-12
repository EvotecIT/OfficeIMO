using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsImportedRelatedPartCaptionFixture() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "imported-related-part-list-of-figures.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing imported related-part list-of-figures fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresImportedRelatedParts.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Header header = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.HeaderParts).Header!;
                Footer footer = Assert.Single(document._wordprocessingDocument.MainDocumentPart.FooterParts).Footer!;

                Assert.Contains(header.Descendants<Text>(), text => text.Text.Contains("Header architecture map", StringComparison.Ordinal));
                Assert.Contains(footer.Descendants<Text>(), text => text.Text.Contains("Footer recovery map", StringComparison.Ordinal));

                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(3, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(
                    new[] {
                        "Figure 1 Body deployment view",
                        "Figure 2 Header architecture map",
                        "Figure 3 Footer recovery map"
                    },
                    report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 1, 1 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Body deployment view", TocText(list));
                Assert.Contains("Figure 2 Header architecture map", TocText(list));
                Assert.Contains("Figure 3 Footer recovery map", TocText(list));
                Assert.Contains(header.Descendants<BookmarkStart>(), bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name));
                Assert.Contains(footer.Descendants<BookmarkStart>(), bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name));
                Assert.Equal(3, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Body deployment view", TocText(list));
                Assert.Contains("Figure 2 Header architecture map", TocText(list));
                Assert.Contains("Figure 3 Footer recovery map", TocText(list));
                Assert.Equal(3, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }

        [Fact]
        public void Test_TableOfContent_RefreshListOfFiguresSupportsImportedNotePartCaptionFixture() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "imported-note-part-list-of-figures.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing imported note-part list-of-figures fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "CaptionListFiguresImportedNoteParts.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Footnote footnote = document._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().First(note => note.Type == null);
                Endnote endnote = document._wordprocessingDocument.MainDocumentPart.EndnotesPart!.Endnotes!.Elements<Endnote>().First(note => note.Type == null);

                Assert.Contains(footnote.Descendants<Text>(), text => text.Text.Contains("Footnote architecture map", StringComparison.Ordinal));
                Assert.Contains(endnote.Descendants<Text>(), text => text.Text.Contains("Endnote recovery map", StringComparison.Ordinal));

                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                WordCaptionListRefreshReport report = list.RefreshListOfFigures();

                Assert.Equal("Figure", report.SequenceIdentifier);
                Assert.Equal(3, report.EntryCount);
                Assert.Equal(0, report.SkippedCaptionCount);
                Assert.Equal(
                    new[] {
                        "Figure 1 Body deployment view",
                        "Figure 2 Footnote architecture map",
                        "Figure 3 Endnote recovery map"
                    },
                    report.Entries.Select(entry => entry.Text).ToArray());
                Assert.Equal(new[] { 1, 1, 1 }, report.Entries.Select(entry => entry.PageNumber).ToArray());
                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Body deployment view", TocText(list));
                Assert.Contains("Figure 2 Footnote architecture map", TocText(list));
                Assert.Contains("Figure 3 Endnote recovery map", TocText(list));
                Assert.Contains(footnote.Descendants<BookmarkStart>(), bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name));
                Assert.Contains(endnote.Descendants<BookmarkStart>(), bookmark => report.Entries.Any(entry => entry.BookmarkName == bookmark.Name));
                Assert.Equal(3, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordTableOfContent list = Assert.IsType<WordTableOfContent>(document.TableOfContent);

                Assert.Contains("List of Figures", TocText(list));
                Assert.Contains("Figure 1 Body deployment view", TocText(list));
                Assert.Contains("Figure 2 Footnote architecture map", TocText(list));
                Assert.Contains("Figure 3 Endnote recovery map", TocText(list));
                Assert.Equal(3, list.SdtBlock.Descendants<Hyperlink>().Count(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Anchor)));
            }
        }
    }
}

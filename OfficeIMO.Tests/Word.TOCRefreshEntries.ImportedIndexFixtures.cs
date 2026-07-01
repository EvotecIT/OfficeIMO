using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_TableOfContent_RefreshIndexSupportsImportedRelatedPartIndexFixture() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "imported-related-part-index.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing imported related-part index fixture: {sourcePath}");

            string filePath = Path.Combine(_directoryWithFiles, "IndexRefreshImportedRelatedParts.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Header header = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.HeaderParts).Header!;
                Footer footer = Assert.Single(document._wordprocessingDocument.MainDocumentPart.FooterParts).Footer!;
                Footnote footnote = document._wordprocessingDocument.MainDocumentPart.FootnotesPart!.Footnotes!.Elements<Footnote>().First(note => note.Type == null);
                Endnote endnote = document._wordprocessingDocument.MainDocumentPart.EndnotesPart!.Endnotes!.Elements<Endnote>().First(note => note.Type == null);

                Assert.Contains(header.Descendants<Text>(), text => text.Text.Contains("Header topic", StringComparison.Ordinal));
                Assert.Contains(footer.Descendants<Text>(), text => text.Text.Contains("Footer topic", StringComparison.Ordinal));
                Assert.Contains(footnote.Descendants<Text>(), text => text.Text.Contains("Footnote topic", StringComparison.Ordinal));
                Assert.Contains(endnote.Descendants<Text>(), text => text.Text.Contains("Endnote topic", StringComparison.Ordinal));

                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                WordIndexRefreshReport report = index.RefreshIndex("Imported Related-Part Index");

                Assert.Equal(5, report.EntryCount);
                Assert.Equal(0, report.SkippedEntryCount);
                Assert.Equal(
                    new[] { "BodyTopic", "EndnoteTopic", "FooterTopic", "FootnoteTopic", "HeaderTopic" },
                    report.Entries.Select(entry => entry.Term).ToArray());
                Assert.All(report.Entries, entry => Assert.Equal(new[] { 1 }, entry.PageNumbers));

                string indexText = TocText(index);
                Assert.Contains("Imported Related-Part Index", indexText);
                Assert.Contains("BodyTopic", indexText);
                Assert.Contains("HeaderTopic", indexText);
                Assert.Contains("FooterTopic", indexText);
                Assert.Contains("FootnoteTopic", indexText);
                Assert.Contains("EndnoteTopic", indexText);
                Assert.True(document.Settings.UpdateFieldsOnOpen);
                Assert.True(document.DocumentIsValid, FormatValidationErrors(document.DocumentValidationErrors));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordTableOfContent index = Assert.IsType<WordTableOfContent>(document.TableOfContent);
                string indexText = TocText(index);

                Assert.Contains("Imported Related-Part Index", indexText);
                Assert.Contains("BodyTopic", indexText);
                Assert.Contains("HeaderTopic", indexText);
                Assert.Contains("FooterTopic", indexText);
                Assert.Contains("FootnoteTopic", indexText);
                Assert.Contains("EndnoteTopic", indexText);
            }
        }
    }
}

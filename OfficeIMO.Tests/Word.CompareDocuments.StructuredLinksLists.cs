using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructureReportsBookmarkHyperlinkAndListChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_links_lists_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Anchor paragraph").AddBookmark("OriginalAnchor");
                document.AddParagraph("Portal: ").AddHyperLink("Open portal", new Uri("https://example.com/source"));
                WordList sourceList = document.AddList(WordListStyle.Numbered);
                sourceList.AddItem("Collect requirements");
                sourceList.AddItem("Draft proposal");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_links_lists_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Anchor paragraph").AddBookmark("UpdatedAnchor");
                document.AddParagraph("Portal: ").AddHyperLink("Open portal", new Uri("https://example.com/target"));
                WordList targetList = document.AddList(WordListStyle.Numbered);
                targetList.AddItem("Collect requirements");
                targetList.AddItem("Draft final proposal");
                document.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding bookmark = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Bookmark &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("bookmark[0]", bookmark.Location);
            Assert.Contains("OriginalAnchor", bookmark.SourceText, StringComparison.Ordinal);
            Assert.Contains("UpdatedAnchor", bookmark.TargetText, StringComparison.Ordinal);

            WordComparisonFinding hyperlink = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Hyperlink &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("hyperlink[0]", hyperlink.Location);
            Assert.Contains("https://example.com/source", hyperlink.SourceText, StringComparison.Ordinal);
            Assert.Contains("https://example.com/target", hyperlink.TargetText, StringComparison.Ordinal);

            WordComparisonFinding list = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.List &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("list[1]", list.Location);
            Assert.Contains("Draft proposal", list.SourceText, StringComparison.Ordinal);
            Assert.Contains("Draft final proposal", list.TargetText, StringComparison.Ordinal);
        }

        [Fact]
        public void CompareStructureOptionsCanDisableBookmarkHyperlinkAndListFindings() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_links_lists_options_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Anchor paragraph").AddBookmark("OriginalAnchor");
                document.AddParagraph("Portal: ").AddHyperLink("Open portal", new Uri("https://example.com/source"));
                document.AddList(WordListStyle.Bulleted).AddItem("Original checklist item");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_links_lists_options_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Anchor paragraph").AddBookmark("UpdatedAnchor");
                document.AddParagraph("Portal: ").AddHyperLink("Open portal", new Uri("https://example.com/target"));
                document.AddList(WordListStyle.Bulleted).AddItem("Updated checklist item");
                document.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareBookmarks = false,
                CompareHyperlinks = false,
                CompareLists = false
            });

            Assert.DoesNotContain(result.Findings, finding => finding.Scope == WordComparisonScope.Bookmark);
            Assert.DoesNotContain(result.Findings, finding => finding.Scope == WordComparisonScope.Hyperlink);
            Assert.DoesNotContain(result.Findings, finding => finding.Scope == WordComparisonScope.List);
        }

        [Fact]
        public void CompareStructureMatchesBookmarkHyperlinkAndListFeaturesAcrossInsertions() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_links_lists_insert_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Anchor paragraph").AddBookmark("StableAnchor");
                document.AddParagraph("Portal: ").AddHyperLink("Open portal", new Uri("https://example.com/stable"));
                WordList sourceList = document.AddList(WordListStyle.Numbered);
                sourceList.AddItem("Collect requirements");
                sourceList.AddItem("Draft proposal");
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_links_lists_insert_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Inserted anchor").AddBookmark("InsertedAnchor");
                document.AddParagraph("Anchor paragraph").AddBookmark("StableAnchor");
                document.AddParagraph("Inserted portal: ").AddHyperLink("Open inserted portal", new Uri("https://example.com/inserted"));
                document.AddParagraph("Portal: ").AddHyperLink("Open portal", new Uri("https://example.com/stable"));
                WordList targetList = document.AddList(WordListStyle.Numbered);
                targetList.AddItem("Inserted task");
                targetList.AddItem("Collect requirements");
                targetList.AddItem("Draft proposal");
                document.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareGeneratedIds = false
            });

            Assert.Single(result.Findings, finding => finding.Scope == WordComparisonScope.Bookmark);
            Assert.Single(result.Findings, finding => finding.Scope == WordComparisonScope.Hyperlink);
            Assert.Single(result.Findings, finding => finding.Scope == WordComparisonScope.List);
            Assert.DoesNotContain(result.Findings, finding =>
                (finding.Scope is WordComparisonScope.Bookmark or WordComparisonScope.Hyperlink or WordComparisonScope.List) &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
        }

        [Fact]
        public void CompareStructureIgnoresListNumberIdsWhenGeneratedIdsAreDisabled() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_list_generated_ids_source.docx");
            CreateRawNumberedParagraphDocument(sourcePath, 101, "Imported list item");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_list_generated_ids_target.docx");
            CreateRawNumberedParagraphDocument(targetPath, 202, "Imported list item");

            WordComparisonResult relaxed = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareGeneratedIds = false
            });

            Assert.DoesNotContain(relaxed.Findings, finding => finding.Scope == WordComparisonScope.List);

            WordComparisonResult strict = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareGeneratedIds = true
            });

            Assert.Contains(strict.Findings, finding =>
                finding.Scope == WordComparisonScope.List &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
        }

        private static void CreateRawNumberedParagraphDocument(string path, int numberingId, string text) {
            using WordDocument document = WordDocument.Create(path);
            document._document.Body!.RemoveAllChildren<Paragraph>();
            document._document.Body!.Append(new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = numberingId })),
                new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
            document.Save(false);
        }
    }
}

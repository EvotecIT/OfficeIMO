using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructureReportsCommentAndRevisionChanges() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_review_source.docx");
            DateTime revisionDate = new DateTime(2026, 6, 28, 12, 0, 0, DateTimeKind.Utc);
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Review target").AddComment("Alice Reviewer", "AR", "Use the older wording.");
                Assert.Single(document.Comments).MarkUnresolved();
                WordParagraph paragraph = document.AddParagraph("Tracked ");
                paragraph.AddInsertedText("Draft", "Alice Reviewer", revisionDate);
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_review_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Review target").AddComment("Alice Reviewer", "AR", "Use the approved wording.");
                Assert.Single(document.Comments).MarkResolved();
                WordParagraph paragraph = document.AddParagraph("Tracked ");
                paragraph.AddInsertedText("Final", "Bob Reviewer", revisionDate);
                document.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            WordComparisonFinding comment = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Comment &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("comment[0]", comment.Location);
            Assert.Contains("Use the older wording.", comment.SourceText, StringComparison.Ordinal);
            Assert.Contains("resolved=false", comment.SourceText, StringComparison.Ordinal);
            Assert.Contains("Use the approved wording.", comment.TargetText, StringComparison.Ordinal);
            Assert.Contains("resolved=true", comment.TargetText, StringComparison.Ordinal);

            WordComparisonFinding revision = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
            Assert.Equal("revision[0]", revision.Location);
            Assert.Contains("author=Alice Reviewer", revision.SourceText, StringComparison.Ordinal);
            Assert.Contains("text=Draft", revision.SourceText, StringComparison.Ordinal);
            Assert.Contains("author=Bob Reviewer", revision.TargetText, StringComparison.Ordinal);
            Assert.Contains("text=Final", revision.TargetText, StringComparison.Ordinal);
        }

        [Fact]
        public void CompareStructureOptionsCanDisableCommentAndRevisionFindings() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_review_options_source.docx");
            DateTime revisionDate = new DateTime(2026, 6, 28, 12, 0, 0, DateTimeKind.Utc);
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Review target").AddComment("Alice Reviewer", "AR", "Source comment.");
                document.AddParagraph("Tracked ").AddDeletedText("Source deletion", "Alice Reviewer", revisionDate);
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_review_options_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Review target").AddComment("Bob Reviewer", "BR", "Target comment.");
                document.AddParagraph("Tracked ").AddDeletedText("Target deletion", "Bob Reviewer", revisionDate);
                document.Save(false);
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareComments = false,
                CompareRevisions = false
            });

            Assert.DoesNotContain(result.Findings, finding => finding.Scope == WordComparisonScope.Comment);
            Assert.DoesNotContain(result.Findings, finding => finding.Scope == WordComparisonScope.Revision);
        }

        [Fact]
        public void CompareStructureOptionsCanIgnoreCommentSubfamilies() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_comment_subfamilies_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Review target").AddComment("Alice Reviewer", "AR", "Source comment text.");
                Assert.Single(document.Comments).MarkUnresolved();
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_comment_subfamilies_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Review target").AddComment("Bob Reviewer", "BR", "Target comment text.");
                WordComment comment = WordComment.GetAllComments(document).Single();
                comment.AddReply("Carol Reviewer", "CR", "Reply text.");
                comment.MarkResolved();
                document.Save(false);
            }

            WordComparisonResult defaultResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            Assert.Contains(defaultResult.Findings, finding => finding.Scope == WordComparisonScope.Comment);

            WordComparisonResult ignoredResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareCommentAuthors = false,
                CompareCommentText = false,
                CompareCommentResolvedState = false,
                CompareCommentReplies = false,
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false
            });

            Assert.DoesNotContain(ignoredResult.Findings, finding => finding.Scope == WordComparisonScope.Comment);
        }

        [Fact]
        public void CompareStructureOptionsCanIgnoreRevisionSubfamilies() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_revision_subfamilies_source.docx");
            DateTime revisionDate = new DateTime(2026, 6, 28, 12, 0, 0, DateTimeKind.Utc);
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Tracked ").AddInsertedText("Draft", "Alice Reviewer", revisionDate);
                document.Save(false);
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_revision_subfamilies_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Tracked ").AddInsertedText("Final", "Bob Reviewer", revisionDate);
                document.Save(false);
            }

            WordComparisonResult defaultResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            Assert.Contains(defaultResult.Findings, finding => finding.Scope == WordComparisonScope.Revision);

            WordComparisonResult ignoredResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, new WordComparisonOptions {
                CompareRevisionAuthors = false,
                CompareRevisionText = false,
                CompareGeneratedIds = false
            });

            Assert.DoesNotContain(ignoredResult.Findings, finding => finding.Scope == WordComparisonScope.Revision);
        }

        [Fact]
        public void CompareStructureReportsWordAuthoredReviewFixture() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_word_authored_review_empty_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("No review metadata here.");
                document.Save(false);
            }

            string targetPath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-comments-revisions.docx"));
            Assert.True(File.Exists(targetPath), $"Missing Word-authored review fixture: {targetPath}");

            var options = new WordComparisonOptions {
                IncludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Comment,
                    WordComparisonScope.Revision
                },
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false
            };

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Comment &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("Word-authored comment for review corpus.", StringComparison.Ordinal) &&
                finding.TargetText.Contains("target=Word-authored comment target", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("Deletion", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Word-authored deletion target", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("Insertion", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Word-authored inserted revision", StringComparison.Ordinal));
        }

        [Fact]
        public void CompareStructureReportsWordAuthoredMoveAndFormattingFixture() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_word_authored_move_formatting_empty_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("No review metadata here.");
                document.Save(false);
            }

            string targetPath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-move-formatting-revisions.docx"));
            Assert.True(File.Exists(targetPath), $"Missing Word-authored move/formatting fixture: {targetPath}");

            var options = new WordComparisonOptions {
                IncludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Revision
                },
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false
            };

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("MoveFrom", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Word-authored move source", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("MoveTo", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Word-authored move source", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("RunFormatting", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Word-authored run formatting target", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("TableFormatting", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Word-authored cell formatting target", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("TableCellFormatting", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Word-authored cell formatting target", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("TableRowFormatting", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Word-authored cell formatting target", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("TableFormatting", StringComparison.Ordinal) &&
                finding.TargetText.Contains("tblGridChange", StringComparison.Ordinal));
        }

        [Fact]
        public void CompareStructureReportsWordAuthoredParagraphAndSectionFormattingFixture() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_word_authored_paragraph_section_empty_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("No review metadata here.");
                document.Save(false);
            }

            string targetPath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-paragraph-section-formatting-revisions.docx"));
            Assert.True(File.Exists(targetPath), $"Missing Word-authored paragraph/section formatting fixture: {targetPath}");

            var options = new WordComparisonOptions {
                IncludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Revision
                },
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false
            };

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("ParagraphFormatting", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Word-authored paragraph formatting target", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText.Contains("SectionFormatting", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Word-authored section formatting anchor", StringComparison.Ordinal));
        }

        [Fact]
        public void CompareStructureReportsWordComAuthoredRelatedPartReviewFixture() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_word_com_related_parts_empty_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("No related-part review metadata here.");
                document.Save(false);
            }

            string targetPath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-related-part-revisions.docx"));
            Assert.True(File.Exists(targetPath), $"Missing Word COM-authored related-part review fixture: {targetPath}");

            var options = new WordComparisonOptions {
                IncludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Revision
                },
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false
            };

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);

            AssertInsertedRelatedPartRevisionFinding(result, "Header", "Insertion", "Word COM header inserted revision");
            AssertInsertedRelatedPartRevisionFinding(result, "Footer", "Insertion", "Word COM footer inserted revision");
            AssertInsertedRelatedPartRevisionFinding(result, "Footnote", "Insertion", "Word COM footnote inserted revision");
            AssertInsertedRelatedPartRevisionFinding(result, "Endnote", "Insertion", "Word COM endnote inserted revision");
            AssertInsertedRelatedPartRevisionFinding(result, "Header", "Deletion", "Word COM header deletion target");
            AssertInsertedRelatedPartRevisionFinding(result, "Footer", "Deletion", "Word COM footer deletion target");
            AssertInsertedRelatedPartRevisionFinding(result, "Footnote", "Deletion", "Word COM footnote deletion target");
            AssertInsertedRelatedPartRevisionFinding(result, "Endnote", "Deletion", "Word COM endnote deletion target");

            string json = result.ToJson();
            Assert.Contains("\"scope\": \"Revision\"", json, StringComparison.Ordinal);
            Assert.Contains("Word COM header inserted revision", json, StringComparison.Ordinal);
            Assert.Contains("word/header", json, StringComparison.OrdinalIgnoreCase);

            string markdown = result.ToMarkdown();
            Assert.Contains("Word COM endnote deletion target", markdown, StringComparison.Ordinal);
            Assert.Contains("Endnote", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void CompareStructureReportsImportedRelatedPartCommentFixture() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_imported_related_part_comments_empty_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("No imported related-part review metadata here.");
                document.Save(false);
            }

            string targetPath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "imported-related-part-comments-revisions.docx"));
            Assert.True(File.Exists(targetPath), $"Missing imported related-part comment fixture: {targetPath}");

            var options = new WordComparisonOptions {
                IncludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Comment,
                    WordComparisonScope.Revision
                },
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false
            };

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);

            AssertInsertedRelatedPartCommentFinding(result, "Header", "Imported corpus header comment target", "Imported corpus header comment.");
            AssertInsertedRelatedPartCommentFinding(result, "Footer", "Imported corpus footer comment target", "Imported corpus footer comment.");
            AssertInsertedRelatedPartCommentFinding(result, "Footnote", "Imported corpus footnote comment target", "Imported corpus footnote comment.");
            AssertInsertedRelatedPartCommentFinding(result, "Endnote", "Imported corpus endnote comment target", "Imported corpus endnote comment.");

            AssertInsertedRelatedPartRevisionFinding(result, "Header", "Insertion", "Imported corpus header insertion");
            AssertInsertedRelatedPartRevisionFinding(result, "Footer", "Insertion", "Imported corpus footer insertion");
            AssertInsertedRelatedPartRevisionFinding(result, "Footnote", "Insertion", "Imported corpus footnote insertion");
            AssertInsertedRelatedPartRevisionFinding(result, "Endnote", "Insertion", "Imported corpus endnote insertion");

            string json = result.ToJson();
            Assert.Contains("\"scope\": \"Comment\"", json, StringComparison.Ordinal);
            Assert.Contains("Imported corpus footnote comment.", json, StringComparison.Ordinal);
            Assert.Contains("word/footnotes", json, StringComparison.OrdinalIgnoreCase);

            string markdown = result.ToMarkdown();
            Assert.Contains("Imported corpus endnote insertion", markdown, StringComparison.Ordinal);
            Assert.Contains("Endnote", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void CompareStructureReturnsImportedReviewMetadataInTargetDocumentOrder() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_imported_review_order_source.docx");
            CreateDocumentWithImportedReviewOrderInputs(sourcePath, "Source", "Alice Reviewer");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_imported_review_order_target.docx");
            CreateDocumentWithImportedReviewOrderInputs(targetPath, "Target", "Bob Reviewer");

            var options = new WordComparisonOptions {
                IncludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Comment,
                    WordComparisonScope.Revision
                },
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false
            };

            WordComparisonResult firstResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);
            WordComparisonResult secondResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);

            string[] firstSequence = firstResult.Findings.Select(FormatReviewFindingSequenceEntry).ToArray();
            string[] secondSequence = secondResult.Findings.Select(FormatReviewFindingSequenceEntry).ToArray();

            Assert.Equal(firstSequence, secondSequence);
            Assert.Equal(new[] {
                "Comment|Modified|comment[2]",
                "Comment|Modified|comment[1]",
                "Revision|Modified|revision[0]",
                "Comment|Modified|comment[0]",
                "Revision|Modified|revision[1]"
            }, firstSequence);

            Assert.Contains(firstResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Comment &&
                finding.Location == "comment[2]" &&
                finding.DetailedLocation.Contains("Body", StringComparison.Ordinal));
            Assert.Contains(firstResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Comment &&
                finding.Location == "comment[1]" &&
                finding.DetailedLocation.Contains("table", StringComparison.Ordinal));
            Assert.Contains(firstResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Comment &&
                finding.Location == "comment[0]" &&
                finding.DetailedLocation.Contains("Header", StringComparison.Ordinal));
            Assert.Contains(firstResult.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.Location == "revision[1]" &&
                finding.DetailedLocation.Contains("Header", StringComparison.Ordinal));
        }

        [Fact]
        public void CompareStructureReportsImportedReviewMetadataInsideContentControlsAndTextBoxes() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_imported_review_containers_source.docx");
            CreateImportedReviewContainerDocument(sourcePath, "Source", "Alice Reviewer");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_imported_review_containers_target.docx");
            CreateImportedReviewContainerDocument(targetPath, "Target", "Bob Reviewer");

            var options = new WordComparisonOptions {
                IncludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Comment,
                    WordComparisonScope.Revision
                },
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false
            };

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Comment &&
                finding.DetailedLocation.Contains("content-control", StringComparison.Ordinal) &&
                finding.SourceText.Contains("Source content-control comment.", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Target content-control comment.", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Comment &&
                finding.DetailedLocation.Contains("text-box", StringComparison.Ordinal) &&
                finding.SourceText.Contains("Source text-box comment.", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Target text-box comment.", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.DetailedLocation.Contains("content-control", StringComparison.Ordinal) &&
                finding.SourceText.Contains("Source content-control insertion", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Target content-control insertion", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.DetailedLocation.Contains("text-box", StringComparison.Ordinal) &&
                finding.SourceText.Contains("Source text-box insertion", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Target text-box insertion", StringComparison.Ordinal));
        }

        [Fact]
        public void CompareStructureReportsImportedMoveAndFormattingRevisions() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_imported_move_formatting_source.docx");
            CreateImportedMoveAndFormattingRevisionDocument(sourcePath, "Source", "Alice Reviewer");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_imported_move_formatting_target.docx");
            CreateImportedMoveAndFormattingRevisionDocument(targetPath, "Target", "Bob Reviewer");

            var options = new WordComparisonOptions {
                IncludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Revision
                },
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false
            };

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.DetailedLocation.Contains("MoveFrom", StringComparison.Ordinal) &&
                finding.SourceText.Contains("Source moved from", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Target moved from", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.DetailedLocation.Contains("MoveTo", StringComparison.Ordinal) &&
                finding.SourceText.Contains("Source moved to", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Target moved to", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.DetailedLocation.Contains("RunFormatting", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.DetailedLocation.Contains("TableFormatting", StringComparison.Ordinal) &&
                finding.DetailedLocation.Contains("table", StringComparison.Ordinal));
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.DetailedLocation.Contains("TableCellFormatting", StringComparison.Ordinal) &&
                finding.DetailedLocation.Contains("table", StringComparison.Ordinal) &&
                finding.SourceText.Contains("Source cell formatting target", StringComparison.Ordinal) &&
                finding.TargetText.Contains("Target cell formatting target", StringComparison.Ordinal));
        }

        [Fact]
        public void CompareStructureReportsImportedReviewMetadataInHeadersFootersAndNotes() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_imported_review_related_parts_source.docx");
            CreateImportedReviewRelatedPartDocument(sourcePath, "Source", "Alice Reviewer");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_imported_review_related_parts_target.docx");
            CreateImportedReviewRelatedPartDocument(targetPath, "Target", "Bob Reviewer");

            var options = new WordComparisonOptions {
                IncludedScopes = new HashSet<WordComparisonScope> {
                    WordComparisonScope.Comment,
                    WordComparisonScope.Revision
                },
                CompareGeneratedIds = false,
                CompareVolatileMetadata = false
            };

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath, options);

            AssertRelatedPartReviewFinding(result, WordComparisonScope.Comment, "Header", "Source header comment.", "Target header comment.");
            AssertRelatedPartReviewFinding(result, WordComparisonScope.Comment, "Footer", "Source footer comment.", "Target footer comment.");
            AssertRelatedPartReviewFinding(result, WordComparisonScope.Comment, "Footnote", "Source footnote comment.", "Target footnote comment.");
            AssertRelatedPartReviewFinding(result, WordComparisonScope.Comment, "Endnote", "Source endnote comment.", "Target endnote comment.");

            AssertRelatedPartReviewFinding(result, WordComparisonScope.Revision, "Header", "Source header insertion", "Target header insertion");
            AssertRelatedPartReviewFinding(result, WordComparisonScope.Revision, "Footer", "Source footer insertion", "Target footer insertion");
            AssertRelatedPartReviewFinding(result, WordComparisonScope.Revision, "Footnote", "Source footnote insertion", "Target footnote insertion");
            AssertRelatedPartReviewFinding(result, WordComparisonScope.Revision, "Endnote", "Source endnote insertion", "Target endnote insertion");
        }

        private static void CreateDocumentWithImportedReviewOrderInputs(string path, string label, string author) {
            DateTime revisionDate = new DateTime(2026, 6, 28, 12, 0, 0, DateTimeKind.Utc);

            using (WordDocument document = WordDocument.Create(path)) {
                document.AddParagraph("Body review target").AddComment(author, "RV", label + " body comment.");

                WordTable table = document.AddTable(1, 1);
                WordParagraph tableParagraph = table.Rows[0].Cells[0].Paragraphs[0];
                tableParagraph.SetText("Table review target ");
                tableParagraph.AddComment(author, "RV", label + " table comment.");
                tableParagraph.AddInsertedText(label + " table revision", author, revisionDate);

                document.AddHeadersAndFooters();
                WordParagraph headerParagraph = document.Header.Default!.AddParagraph("Header review target ");
                headerParagraph.AddComment(author, "RV", label + " header comment.");
                headerParagraph.AddInsertedText(label + " header revision", author, revisionDate);
                document.Save(false);
            }

            ReverseImportedCommentStorageOrder(path);
        }

        private static void ReverseImportedCommentStorageOrder(string path) {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
            Comments? comments = mainPart.WordprocessingCommentsPart?.Comments;
            if (comments != null) {
                Comment[] reversedComments = comments.Elements<Comment>()
                    .Reverse()
                    .Select(comment => (Comment)comment.CloneNode(true))
                    .ToArray();
                comments.RemoveAllChildren<Comment>();
                comments.Append(reversedComments);
                comments.Save();
            }

            W15.CommentsEx? commentsEx = mainPart.WordprocessingCommentsExPart?.CommentsEx;
            if (commentsEx != null) {
                W15.CommentEx[] reversedCommentsEx = commentsEx.Elements<W15.CommentEx>()
                    .Reverse()
                    .Select(comment => (W15.CommentEx)comment.CloneNode(true))
                    .ToArray();
                commentsEx.RemoveAllChildren<W15.CommentEx>();
                commentsEx.Append(reversedCommentsEx);
                commentsEx.Save();
            }
        }

        private static string FormatReviewFindingSequenceEntry(WordComparisonFinding finding) {
            return string.Join("|", finding.Scope, finding.ChangeKind, finding.Location);
        }

        private static void AssertRelatedPartReviewFinding(
            WordComparisonResult result,
            WordComparisonScope scope,
            string location,
            string sourceText,
            string targetText) {
            Assert.Contains(result.Findings, finding =>
                finding.Scope == scope &&
                finding.DetailedLocation.Contains(location, StringComparison.Ordinal) &&
                finding.SourceText.Contains(sourceText, StringComparison.Ordinal) &&
                finding.TargetText.Contains(targetText, StringComparison.Ordinal));
        }

        private static void AssertInsertedRelatedPartRevisionFinding(
            WordComparisonResult result,
            string location,
            string revisionType,
            string targetText) {
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Revision &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.DetailedLocation.Contains(location, StringComparison.Ordinal) &&
                finding.DetailedLocation.Contains("revision", StringComparison.OrdinalIgnoreCase) &&
                finding.TargetText.Contains(revisionType, StringComparison.Ordinal) &&
                finding.TargetText.Contains(targetText, StringComparison.Ordinal));
        }

        private static void AssertInsertedRelatedPartCommentFinding(
            WordComparisonResult result,
            string location,
            string targetText,
            string commentText) {
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Comment &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.DetailedLocation.Contains(location, StringComparison.Ordinal) &&
                finding.DetailedLocation.Contains("comment", StringComparison.OrdinalIgnoreCase) &&
                finding.TargetText.Contains(targetText, StringComparison.Ordinal) &&
                finding.TargetText.Contains(commentText, StringComparison.Ordinal));
        }
    }
}

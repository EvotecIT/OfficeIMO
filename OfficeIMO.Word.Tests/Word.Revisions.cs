using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AcceptRevisions_RemovesTrackedChanges() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChanges.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Before");

                var paragraph = document.AddParagraph();
                paragraph.AddInsertedText("Added", "Codex");
                paragraph.AddDeletedText("Removed", "Codex");

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.NotNull(document._document);
                var body = document._document!.Body;
                Assert.NotNull(body);
                Assert.Contains(body!.Descendants<InsertedRun>(), run => run.InnerText == "Added");
                Assert.Contains(body.Descendants<DeletedRun>(), run => run.InnerText == "Removed");

                document.AcceptRevisions();

                Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText == "Added");
                Assert.DoesNotContain(body.Descendants<DeletedRun>(), run => run.InnerText == "Removed");
                Assert.Contains(document.Paragraphs, p => p.Text == "Before");
                Assert.Contains(document.Paragraphs, p => p.Text == "Added");
            }
        }

        [Fact]
        public void Test_RejectRevisions_RemovesInsertions() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesReject.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddInsertedText("Added", "Codex");
                paragraph.AddDeletedText("Removed", "Codex");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.RejectRevisions();
                Assert.NotNull(document._document);
                var body = document._document!.Body;
                Assert.NotNull(body);
                Assert.DoesNotContain(body!.Descendants<InsertedRun>(), run => run.InnerText == "Added");
                Assert.DoesNotContain(body.Descendants<DeletedRun>(), run => run.InnerText == "Removed");
                Assert.Contains(document.Paragraphs, p => p.Text == "Removed");
            }
        }

        [Fact]
        public void Test_TrackedChanges_Validation() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesValidation.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddInsertedText("Added", "Codex");
                paragraph.AddDeletedText("Removed", "Codex");
                document.Save();

                var errors = document.ValidateDocument();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_ConvertRevisionsToMarkup_PreservesTextWithFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesMarkup.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddInsertedText("Added", "Codex");
                paragraph.AddDeletedText("Removed", "Codex");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.ConvertRevisionsToMarkup();

                Assert.NotNull(document._document);
                var body = document._document!.Body;
                Assert.NotNull(body);
                Assert.DoesNotContain(body!.Descendants<InsertedRun>(), r => r.InnerText == "Added");
                Assert.DoesNotContain(body.Descendants<DeletedRun>(), r => r.InnerText == "Removed");

                var insertedRun = body.Descendants<Run>().FirstOrDefault(r => r.InnerText == "Added");
                Assert.NotNull(insertedRun);
                Assert.NotNull(insertedRun!.RunProperties);
                Assert.NotNull(insertedRun.RunProperties!.Underline);
                Assert.Equal("0000FF", insertedRun.RunProperties.Color?.Val);

                var deletedRun = body.Descendants<Run>().FirstOrDefault(r => r.InnerText == "Removed");
                Assert.NotNull(deletedRun);
                Assert.Contains(deletedRun!.Descendants<Text>(), text => text.Text == "Removed");
                Assert.Empty(deletedRun.Descendants<DeletedText>());
                Assert.NotNull(deletedRun!.RunProperties);
                Assert.NotNull(deletedRun.RunProperties!.Strike);
                Assert.Equal("FF0000", deletedRun.RunProperties.Color?.Val);
            }
        }

        [Fact]
        public void Test_AcceptRevisions_HandlesInsertionWithoutRuns() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesEmptyInsertion.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph._paragraph.Append(new Inserted() { Id = "1", Author = "Codex" });
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AcceptRevisions();
                Assert.NotNull(document._document);
                var body = document._document.Body;
                Assert.NotNull(body);
                Assert.Empty(body.Descendants<Inserted>());
            }
        }

        [Fact]
        public void Test_AcceptRevisions_ByAuthor_OnlyAcceptsMatchingChanges() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesAcceptByAuthor.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddInsertedText("AddedByAlice", "Alice");
                paragraph.AddInsertedText("AddedByBob", "Bob");
                paragraph.AddDeletedText("RemovedByAlice", "Alice");
                paragraph.AddDeletedText("RemovedByBob", "Bob");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AcceptRevisions("Alice");

                Assert.NotNull(document._document);
                var body = document._document!.Body;
                Assert.NotNull(body);

                Assert.DoesNotContain(body!.Descendants<InsertedRun>(), run => run.Author?.Value == "Alice");
                Assert.DoesNotContain(body.Descendants<DeletedRun>(), run => run.Author?.Value == "Alice");
                Assert.Contains(body.Descendants<InsertedRun>(), run => run.Author?.Value == "Bob" && run.InnerText == "AddedByBob");
                Assert.Contains(body.Descendants<DeletedRun>(), run => run.Author?.Value == "Bob" && run.InnerText == "RemovedByBob");
                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "AddedByAlice");
            }
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void Test_AcceptRejectRevisions_ByAuthor_RejectsBlankAuthor(string? authorName) {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesBlankAuthor.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph().AddInsertedText("AddedByAlice", "Alice");

                Assert.Throws<ArgumentException>(() => document.AcceptRevisions(authorName!));
                Assert.Throws<ArgumentException>(() => document.RejectRevisions(authorName!));
                Assert.Contains(document._document.Body!.Descendants<InsertedRun>(), run => run.Author?.Value == "Alice");
            }
        }

        [Fact]
        public void Test_AcceptRevisions_ByAuthor_PreservesUnmatchedNestedRevisions() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesAcceptNestedByAuthor.docx");
            File.Delete(filePath);

            DateTime revisionDate = new DateTime(2026, 6, 30, 12, 0, 0, DateTimeKind.Utc);
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph._paragraph.Append(new InsertedRun(
                    new Run(new Text("AddedByAlice")),
                    new InsertedRun(new Run(new Text("NestedByBob"))) {
                        Author = "Bob",
                        Date = revisionDate,
                        Id = "9102"
                    }) {
                    Author = "Alice",
                    Date = revisionDate,
                    Id = "9101"
                });
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport report = document.AcceptRevisions(new WordRevisionFilter { Author = "Alice" });

                Assert.Single(report.MatchedRevisions);
                Body body = document._document.Body!;
                Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.Author?.Value == "Alice");
                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "AddedByAlice");
                Assert.Contains(body.Descendants<InsertedRun>(), run => run.Author?.Value == "Bob" && run.InnerText == "NestedByBob");
            }
        }

        [Fact]
        public void Test_RejectRevisions_ByAuthor_OnlyRejectsMatchingChanges() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesRejectByAuthor.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddInsertedText("AddedByAlice", "Alice");
                paragraph.AddInsertedText("AddedByBob", "Bob");
                paragraph.AddDeletedText("RemovedByAlice", "Alice");
                paragraph.AddDeletedText("RemovedByBob", "Bob");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.RejectRevisions("Alice");

                Assert.NotNull(document._document);
                var body = document._document!.Body;
                Assert.NotNull(body);

                Assert.DoesNotContain(body!.Descendants<InsertedRun>(), run => run.Author?.Value == "Alice");
                Assert.DoesNotContain(body.Descendants<DeletedRun>(), run => run.Author?.Value == "Alice");
                Assert.Contains(body.Descendants<InsertedRun>(), run => run.Author?.Value == "Bob" && run.InnerText == "AddedByBob");
                Assert.Contains(body.Descendants<DeletedRun>(), run => run.Author?.Value == "Bob" && run.InnerText == "RemovedByBob");
                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "RemovedByAlice");
                Assert.DoesNotContain(body.Descendants<Run>(), run => run.InnerText == "AddedByAlice");
            }
        }

        [Fact]
        public void Test_AcceptRevisions_FilterByAuthorTypeAndDate_OnlyAcceptsMatchingChanges() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesAcceptScoped.docx");
            File.Delete(filePath);
            DateTime oldDate = new DateTime(2026, 1, 1, 8, 0, 0, DateTimeKind.Utc);
            DateTime recentDate = new DateTime(2026, 6, 28, 8, 0, 0, DateTimeKind.Utc);

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddInsertedText("OldAlice", "Alice", oldDate);
                paragraph.AddInsertedText("RecentAlice", "Alice", recentDate);
                paragraph.AddInsertedText("RecentBob", "Bob", recentDate);
                paragraph.AddDeletedText("RemovedAlice", "Alice", recentDate);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport report = document.AcceptRevisions(new WordRevisionFilter {
                    Author = "Alice",
                    RevisionType = WordReviewRevisionType.Insertion,
                    DateFrom = new DateTime(2026, 6, 1, 0, 0, 0, DateTimeKind.Utc)
                });

                Assert.Equal(WordRevisionOperationKind.Accept, report.Operation);
                WordRevisionInfo matched = Assert.Single(report.MatchedRevisions);
                Assert.Equal("RecentAlice", matched.AffectedText);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                Body body = document._document.Body!;

                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "RecentAlice");
                Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText == "RecentAlice");
                Assert.Contains(body.Descendants<InsertedRun>(), run => run.InnerText == "OldAlice");
                Assert.Contains(body.Descendants<InsertedRun>(), run => run.InnerText == "RecentBob");
                Assert.Contains(body.Descendants<DeletedRun>(), run => run.InnerText == "RemovedAlice");
            }
        }

        [Fact]
        public void Test_RejectRevisions_FilterByExplicitId_OnlyRejectsMatchingRevision() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesRejectById.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddDeletedText("RestoreThis", "Alice");
                paragraph.AddDeletedText("KeepDeleted", "Alice");
                document.Save();
            }

            string revisionId;
            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                revisionId = Assert.Single(document.InspectReview().Revisions, revision => revision.AffectedText == "RestoreThis").Id!;
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport report = document.RejectRevisions(new WordRevisionFilter { RevisionId = revisionId });

                WordRevisionInfo matched = Assert.Single(report.MatchedRevisions);
                Assert.Equal("RestoreThis", matched.AffectedText);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                Body body = document._document.Body!;

                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "RestoreThis");
                Assert.DoesNotContain(body.Descendants<DeletedRun>(), run => run.InnerText == "RestoreThis");
                Assert.Contains(body.Descendants<DeletedRun>(), run => run.InnerText == "KeepDeleted");
            }
        }

        [Fact]
        public void Test_Revisions_ParagraphAndTableScopedOperationsOnlyAffectMatchingScope() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesParagraphAndTableScope.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph first = document.AddParagraph();
                first.AddInsertedText("ParagraphOne", "Alice");
                WordParagraph second = document.AddParagraph();
                second.AddInsertedText("ParagraphTwo", "Alice");

                WordTable table = document.AddTable(1, 1);
                table.FirstRow.FirstCell.Paragraphs[0].AddInsertedText("TableInsertion", "Alice");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordParagraph first = document.Paragraphs[0];
                WordRevisionOperationReport paragraphReport = document.AcceptRevisionsInParagraph(first, new WordRevisionFilter {
                    RevisionType = WordReviewRevisionType.Insertion
                });
                Assert.Single(paragraphReport.MatchedRevisions);

                WordRevisionOperationReport tableReport = document.RejectRevisions(new WordRevisionFilter {
                    RevisionType = WordReviewRevisionType.Insertion,
                    IsInTable = true
                });
                Assert.Single(tableReport.MatchedRevisions);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                Body body = document._document.Body!;

                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "ParagraphOne");
                Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText == "ParagraphOne");
                Assert.Contains(body.Descendants<InsertedRun>(), run => run.InnerText == "ParagraphTwo");
                Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText == "TableInsertion");
                Assert.DoesNotContain(body.Descendants<Run>(), run => run.InnerText == "TableInsertion");
            }
        }

        [Fact]
        public void Test_Revisions_TableScopedAcceptFinalizesNestedPromotedRevisions() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesNestedTableScope.docx");
            File.Delete(filePath);
            DateTime revisionDate = new DateTime(2026, 7, 2, 8, 0, 0, DateTimeKind.Utc);

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(1, 1);
                Paragraph paragraph = table.FirstRow.FirstCell.Paragraphs[0]._paragraph;
                paragraph.RemoveAllChildren<Run>();
                paragraph.Append(new InsertedRun(
                    new Run(new Text("Outer ") { Space = SpaceProcessingModeValues.Preserve }),
                    new InsertedRun(new Run(new Text("Inner"))) {
                        Id = "9102",
                        Author = "Alice",
                        Date = revisionDate
                    }) {
                    Id = "9101",
                    Author = "Alice",
                    Date = revisionDate
                });
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport report = document.AcceptRevisions(new WordRevisionFilter {
                    RevisionType = WordReviewRevisionType.Insertion,
                    IsInTable = true
                });

                Assert.Equal(2, report.MatchedCount);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                Body body = document._document.Body!;

                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "Outer ");
                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "Inner");
                Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText.Contains("Inner", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_Revisions_MoveRangeOperationsPromoteMatchingMoveText() {
            string acceptPath = Path.Combine(_directoryWithFiles, "TrackedChangesAcceptMoveTo.docx");
            CreateImportedMoveAndFormattingRevisionDocument(acceptPath, "Accept", "Alice");

            using (WordDocument document = WordDocument.Load(acceptPath)) {
                WordRevisionOperationReport report = document.AcceptRevisions(new WordRevisionFilter {
                    RevisionType = WordReviewRevisionType.MoveTo
                });

                WordRevisionInfo matched = Assert.Single(report.MatchedRevisions);
                Assert.Equal(WordReviewRevisionType.MoveTo, matched.RevisionType);
                Assert.Equal("Accept moved to", matched.AffectedText);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(acceptPath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                Body body = document._document.Body!;

                Assert.DoesNotContain(body.Descendants<MoveToRun>(), run => run.InnerText == "Accept moved to");
                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "Accept moved to");
                Assert.Contains(body.Descendants<MoveFromRun>(), run => run.InnerText == "Accept moved from");
            }

            string rejectPath = Path.Combine(_directoryWithFiles, "TrackedChangesRejectMoveFrom.docx");
            CreateImportedMoveAndFormattingRevisionDocument(rejectPath, "Reject", "Alice");

            using (WordDocument document = WordDocument.Load(rejectPath)) {
                WordRevisionOperationReport report = document.RejectRevisions(new WordRevisionFilter {
                    RevisionType = WordReviewRevisionType.MoveFrom
                });

                WordRevisionInfo matched = Assert.Single(report.MatchedRevisions);
                Assert.Equal(WordReviewRevisionType.MoveFrom, matched.RevisionType);
                Assert.Equal("Reject moved from", matched.AffectedText);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(rejectPath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                Body body = document._document.Body!;

                Assert.DoesNotContain(body.Descendants<MoveFromRun>(), run => run.InnerText == "Reject moved from");
                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "Reject moved from");
                Assert.Contains(body.Descendants<MoveToRun>(), run => run.InnerText == "Reject moved to");
            }
        }

        [Fact]
        public void Test_Revisions_WordAuthoredInsertDeleteFixtureSupportsScopedOperations() {
            string filePath = CopyFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-comments-revisions.docx"));

            string insertionId;
            string deletionId;
            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();
                insertionId = Assert.Single(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.AffectedText == "Word-authored inserted revision").Id!;
                deletionId = Assert.Single(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.AffectedText == "Word-authored deletion target").Id!;
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport acceptReport = document.AcceptRevisions(new WordRevisionFilter {
                    RevisionId = insertionId
                });
                WordRevisionInfo accepted = Assert.Single(acceptReport.MatchedRevisions);
                Assert.Equal(WordReviewRevisionType.Insertion, accepted.RevisionType);
                Assert.Equal("Word-authored inserted revision", accepted.AffectedText);

                WordRevisionOperationReport rejectReport = document.RejectRevisions(new WordRevisionFilter {
                    RevisionId = deletionId
                });
                WordRevisionInfo rejected = Assert.Single(rejectReport.MatchedRevisions);
                Assert.Equal(WordReviewRevisionType.Deletion, rejected.RevisionType);
                Assert.Equal("Word-authored deletion target", rejected.AffectedText);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();
                Assert.DoesNotContain(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.AffectedText == "Word-authored inserted revision");
                Assert.DoesNotContain(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.AffectedText == "Word-authored deletion target");

                Body body = document._document.Body!;
                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "Word-authored inserted revision");
                Assert.Contains(body.Descendants<Text>(), text => text.Text == "Word-authored deletion target");
                Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText == "Word-authored inserted revision");
                Assert.DoesNotContain(body.Descendants<DeletedRun>(), run => run.InnerText == "Word-authored deletion target");
                Assert.DoesNotContain(body.Descendants<DeletedText>(), text => text.Text == "Word-authored deletion target");
            }
        }

        [Fact]
        public void Test_Revisions_WordComAuthoredFixtureSupportsScopedOperations() {
            string filePath = CopyFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-body-table-comments-revisions.docx"));

            string insertionId;
            string deletionId;
            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();
                insertionId = Assert.Single(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.AffectedText == "Word COM inserted body revision").Id!;
                deletionId = Assert.Single(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.AffectedText == "Word COM deletion target").Id!;
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport acceptReport = document.AcceptRevisions(new WordRevisionFilter {
                    RevisionId = insertionId
                });
                WordRevisionInfo accepted = Assert.Single(acceptReport.MatchedRevisions);
                Assert.Equal(WordReviewRevisionType.Insertion, accepted.RevisionType);
                Assert.Equal("Word COM inserted body revision", accepted.AffectedText);

                WordRevisionOperationReport rejectReport = document.RejectRevisions(new WordRevisionFilter {
                    RevisionId = deletionId
                });
                WordRevisionInfo rejected = Assert.Single(rejectReport.MatchedRevisions);
                Assert.Equal(WordReviewRevisionType.Deletion, rejected.RevisionType);
                Assert.Equal("Word COM deletion target", rejected.AffectedText);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();
                Assert.DoesNotContain(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.AffectedText == "Word COM inserted body revision");
                Assert.DoesNotContain(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.AffectedText == "Word COM deletion target");
                Assert.Equal(2, review.CommentCount);

                Body body = document._document.Body!;
                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "Word COM inserted body revision");
                Assert.Contains(body.Descendants<Text>(), text => text.Text == "Word COM deletion target");
                Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText == "Word COM inserted body revision");
                Assert.DoesNotContain(body.Descendants<DeletedRun>(), run => run.InnerText == "Word COM deletion target");
                Assert.DoesNotContain(body.Descendants<DeletedText>(), text => text.Text == "Word COM deletion target");
            }
        }

        [Fact]
        public void Test_Revisions_WordComAuthoredRelatedPartFixtureSupportsScopedOperations() {
            string filePath = CopyFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-related-part-revisions.docx"));

            string headerInsertionId;
            string footerDeletionId;
            string footnoteInsertionId;
            string endnoteDeletionId;
            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();
                headerInsertionId = Assert.Single(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.LocationKind == WordReviewLocationKind.Header &&
                    revision.AffectedText == "Word COM header inserted revision").Id!;
                footerDeletionId = Assert.Single(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.LocationKind == WordReviewLocationKind.Footer &&
                    revision.AffectedText == "Word COM footer deletion target").Id!;
                footnoteInsertionId = Assert.Single(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.LocationKind == WordReviewLocationKind.Footnote &&
                    revision.AffectedText == "Word COM footnote inserted revision").Id!;
                endnoteDeletionId = Assert.Single(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.LocationKind == WordReviewLocationKind.Endnote &&
                    revision.AffectedText == "Word COM endnote deletion target").Id!;
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport headerReport = document.AcceptRevisions(new WordRevisionFilter {
                    RevisionId = headerInsertionId,
                    LocationKind = WordReviewLocationKind.Header
                });

                WordRevisionInfo headerRevision = Assert.Single(headerReport.MatchedRevisions);
                Assert.Equal("Word COM header inserted revision", headerRevision.AffectedText);

                WordRevisionOperationReport footerReport = document.RejectRevisions(new WordRevisionFilter {
                    RevisionId = footerDeletionId,
                    LocationKind = WordReviewLocationKind.Footer
                });

                WordRevisionInfo footerRevision = Assert.Single(footerReport.MatchedRevisions);
                Assert.Equal("Word COM footer deletion target", footerRevision.AffectedText);

                WordRevisionOperationReport footnoteReport = document.AcceptRevisions(new WordRevisionFilter {
                    RevisionId = footnoteInsertionId,
                    LocationKind = WordReviewLocationKind.Footnote
                });

                WordRevisionInfo footnoteRevision = Assert.Single(footnoteReport.MatchedRevisions);
                Assert.Equal("Word COM footnote inserted revision", footnoteRevision.AffectedText);

                WordRevisionOperationReport endnoteReport = document.RejectRevisions(new WordRevisionFilter {
                    RevisionId = endnoteDeletionId,
                    LocationKind = WordReviewLocationKind.Endnote
                });

                WordRevisionInfo endnoteRevision = Assert.Single(endnoteReport.MatchedRevisions);
                Assert.Equal("Word COM endnote deletion target", endnoteRevision.AffectedText);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();
                Assert.DoesNotContain(review.Revisions, revision =>
                    revision.LocationKind == WordReviewLocationKind.Header &&
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.AffectedText == "Word COM header inserted revision");
                Assert.DoesNotContain(review.Revisions, revision =>
                    revision.LocationKind == WordReviewLocationKind.Footer &&
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.AffectedText == "Word COM footer deletion target");
                Assert.DoesNotContain(review.Revisions, revision =>
                    revision.LocationKind == WordReviewLocationKind.Footnote &&
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.AffectedText == "Word COM footnote inserted revision");
                Assert.DoesNotContain(review.Revisions, revision =>
                    revision.LocationKind == WordReviewLocationKind.Endnote &&
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.AffectedText == "Word COM endnote deletion target");

                Assert.Contains(review.Revisions, revision =>
                    revision.LocationKind == WordReviewLocationKind.Header &&
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.AffectedText == "Word COM header deletion target");
                Assert.Contains(review.Revisions, revision =>
                    revision.LocationKind == WordReviewLocationKind.Footer &&
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.AffectedText == "Word COM footer inserted revision");
                Assert.Contains(review.Revisions, revision =>
                    revision.LocationKind == WordReviewLocationKind.Footnote &&
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.AffectedText == "Word COM footnote deletion target");
                Assert.Contains(review.Revisions, revision =>
                    revision.LocationKind == WordReviewLocationKind.Endnote &&
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.AffectedText == "Word COM endnote inserted revision");
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;

                Assert.Contains(mainPart.HeaderParts.SelectMany(part => part.Header!.Descendants<Run>()), run => run.InnerText == "Word COM header inserted revision");
                Assert.DoesNotContain(mainPart.HeaderParts.SelectMany(part => part.Header!.Descendants<InsertedRun>()), run => run.InnerText == "Word COM header inserted revision");
                Assert.Contains(mainPart.HeaderParts.SelectMany(part => part.Header!.Descendants<DeletedRun>()), run => run.InnerText == "Word COM header deletion target");

                Assert.Contains(mainPart.FooterParts.SelectMany(part => part.Footer!.Descendants<Text>()), text => text.Text == "Word COM footer deletion target");
                Assert.DoesNotContain(mainPart.FooterParts.SelectMany(part => part.Footer!.Descendants<DeletedRun>()), run => run.InnerText == "Word COM footer deletion target");
                Assert.Contains(mainPart.FooterParts.SelectMany(part => part.Footer!.Descendants<InsertedRun>()), run => run.InnerText == "Word COM footer inserted revision");

                Assert.Contains(mainPart.FootnotesPart!.Footnotes!.Descendants<Run>(), run => run.InnerText == "Word COM footnote inserted revision");
                Assert.DoesNotContain(mainPart.FootnotesPart.Footnotes.Descendants<InsertedRun>(), run => run.InnerText == "Word COM footnote inserted revision");
                Assert.Contains(mainPart.FootnotesPart.Footnotes.Descendants<DeletedRun>(), run => run.InnerText == "Word COM footnote deletion target");

                Assert.Contains(mainPart.EndnotesPart!.Endnotes!.Descendants<Text>(), text => text.Text == "Word COM endnote deletion target");
                Assert.DoesNotContain(mainPart.EndnotesPart.Endnotes.Descendants<DeletedRun>(), run => run.InnerText == "Word COM endnote deletion target");
                Assert.Contains(mainPart.EndnotesPart.Endnotes.Descendants<InsertedRun>(), run => run.InnerText == "Word COM endnote inserted revision");
            }
        }

        [Fact]
        public void Test_Revisions_WordAuthoredMoveFixtureSupportsScopedOperations() {
            string filePath = CopyFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-move-formatting-revisions.docx"));

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport moveToReport = document.AcceptRevisions(new WordRevisionFilter {
                    RevisionType = WordReviewRevisionType.MoveTo
                });
                WordRevisionInfo moveTo = Assert.Single(moveToReport.MatchedRevisions);
                Assert.Equal("Word-authored move source", moveTo.AffectedText);

                WordRevisionOperationReport moveFromReport = document.RejectRevisions(new WordRevisionFilter {
                    RevisionType = WordReviewRevisionType.MoveFrom
                });
                WordRevisionInfo moveFrom = Assert.Single(moveFromReport.MatchedRevisions);
                Assert.Equal("Word-authored move source", moveFrom.AffectedText);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                Body body = document._document.Body!;
                Assert.DoesNotContain(body.Descendants<MoveToRun>(), run => run.InnerText == "Word-authored move source");
                Assert.DoesNotContain(body.Descendants<MoveFromRun>(), run => run.InnerText == "Word-authored move source");
                Assert.True(body.Descendants<Run>().Count(run => run.InnerText == "Word-authored move source") >= 2);

                WordReviewInfo review = document.InspectReview();
                Assert.DoesNotContain(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.MoveTo);
                Assert.DoesNotContain(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.MoveFrom);
                Assert.Contains(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.RunFormatting);
                Assert.Contains(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.TableFormatting);
                Assert.Contains(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.TableRowFormatting);
                Assert.Contains(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.TableCellFormatting);
            }
        }

        [Fact]
        public void Test_Revisions_WordAuthoredParagraphSectionFormattingFixtureSupportsScopedOperations() {
            string filePath = CopyFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-paragraph-section-formatting-revisions.docx"));

            string paragraphRevisionId;
            string sectionRevisionId;
            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();
                paragraphRevisionId = Assert.Single(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.ParagraphFormatting &&
                    revision.LocationText.Contains("Word-authored paragraph formatting target", StringComparison.Ordinal)).Id!;
                sectionRevisionId = Assert.Single(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.SectionFormatting &&
                    revision.LocationText.Contains("Word-authored section formatting anchor", StringComparison.Ordinal)).Id!;
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport paragraphReport = document.AcceptRevisions(new WordRevisionFilter {
                    RevisionId = paragraphRevisionId,
                    RevisionType = WordReviewRevisionType.ParagraphFormatting
                });

                WordRevisionInfo paragraphRevision = Assert.Single(paragraphReport.MatchedRevisions);
                Assert.Equal(WordReviewRevisionType.ParagraphFormatting, paragraphRevision.RevisionType);

                WordRevisionOperationReport sectionReport = document.RejectRevisions(new WordRevisionFilter {
                    RevisionId = sectionRevisionId,
                    RevisionType = WordReviewRevisionType.SectionFormatting
                });

                WordRevisionInfo sectionRevision = Assert.Single(sectionReport.MatchedRevisions);
                Assert.Equal(WordReviewRevisionType.SectionFormatting, sectionRevision.RevisionType);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();
                Assert.DoesNotContain(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.ParagraphFormatting);
                Assert.DoesNotContain(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.SectionFormatting);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                Body body = wordDocument.MainDocumentPart!.Document.Body!;

                Paragraph paragraph = Assert.Single(body.Descendants<Paragraph>(), item => item.InnerText == "Word-authored paragraph formatting target");
                ParagraphProperties paragraphProperties = paragraph.ParagraphProperties!;
                Assert.Equal("720", paragraphProperties.Indentation!.Left!.Value);
                Assert.Equal(JustificationValues.Center, paragraphProperties.Justification!.Val!.Value);
                Assert.Null(paragraphProperties.ParagraphPropertiesChange);

                SectionProperties sectionProperties = body.Elements<SectionProperties>().Single();
                Assert.Null(sectionProperties.GetFirstChild<PageSize>());
                Assert.Equal(1440U, sectionProperties.GetFirstChild<PageMargin>()!.Left!.Value);
                Assert.Null(sectionProperties.GetFirstChild<SectionPropertiesChange>());
            }
        }

        [Fact]
        public void Test_Revisions_FormattingOperationsAcceptOrRejectStoredProperties() {
            string acceptPath = Path.Combine(_directoryWithFiles, "TrackedChangesAcceptFormatting.docx");
            CreateImportedFormattingRevisionDocument(acceptPath, "Accept", "Alice");

            using (WordDocument document = WordDocument.Load(acceptPath)) {
                WordRevisionOperationReport report = document.AcceptRevisions(new WordRevisionFilter {
                    Author = "Alice"
                });

                Assert.Equal(6, report.MatchedCount);
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(acceptPath, false)) {
                AssertAcceptedFormatting(wordDocument);
            }

            string rejectPath = Path.Combine(_directoryWithFiles, "TrackedChangesRejectFormatting.docx");
            CreateImportedFormattingRevisionDocument(rejectPath, "Reject", "Alice");

            using (WordDocument document = WordDocument.Load(rejectPath)) {
                WordRevisionOperationReport report = document.RejectRevisions(new WordRevisionFilter {
                    Author = "Alice"
                });

                Assert.Equal(6, report.MatchedCount);
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(rejectPath, false)) {
                AssertRejectedFormatting(wordDocument);
            }
        }

        [Fact]
        public void Test_Revisions_LocationScopedOperationsAffectHeadersFootersAndNotes() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesRelatedPartScopes.docx");
            CreateImportedReviewRelatedPartDocument(filePath, "Scoped", "Alice");

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport footerReport = document.AcceptRevisions(new WordRevisionFilter {
                    RevisionType = WordReviewRevisionType.Insertion,
                    LocationKind = WordReviewLocationKind.Footer
                });

                WordRevisionInfo footerRevision = Assert.Single(footerReport.MatchedRevisions);
                Assert.Equal(WordReviewLocationKind.Footer, footerRevision.LocationKind);
                Assert.Equal("Scoped footer insertion", footerRevision.AffectedText);

                WordRevisionOperationReport footnoteReport = document.RejectRevisions(new WordRevisionFilter {
                    RevisionType = WordReviewRevisionType.Insertion,
                    LocationKind = WordReviewLocationKind.Footnote
                });

                WordRevisionInfo footnoteRevision = Assert.Single(footnoteReport.MatchedRevisions);
                Assert.Equal(WordReviewLocationKind.Footnote, footnoteRevision.LocationKind);
                Assert.Equal("Scoped footnote insertion", footnoteRevision.AffectedText);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.DoesNotContain(review.Revisions, revision => revision.LocationKind == WordReviewLocationKind.Footer);
                Assert.DoesNotContain(review.Revisions, revision => revision.LocationKind == WordReviewLocationKind.Footnote);
                Assert.Contains(review.Revisions, revision => revision.LocationKind == WordReviewLocationKind.Header && revision.AffectedText == "Scoped header insertion");
                Assert.Contains(review.Revisions, revision => revision.LocationKind == WordReviewLocationKind.Endnote && revision.AffectedText == "Scoped endnote insertion");
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;

                Assert.Contains(mainPart.FooterParts.SelectMany(part => part.Footer!.Descendants<Run>()), run => run.InnerText == "Scoped footer insertion");
                Assert.DoesNotContain(mainPart.FooterParts.SelectMany(part => part.Footer!.Descendants<InsertedRun>()), run => run.InnerText == "Scoped footer insertion");
                Assert.DoesNotContain(mainPart.FootnotesPart!.Footnotes!.Descendants<Run>(), run => run.InnerText == "Scoped footnote insertion");
                Assert.Contains(mainPart.HeaderParts.SelectMany(part => part.Header!.Descendants<InsertedRun>()), run => run.InnerText == "Scoped header insertion");
                Assert.Contains(mainPart.EndnotesPart!.Endnotes!.Descendants<InsertedRun>(), run => run.InnerText == "Scoped endnote insertion");
            }
        }

        [Fact]
        public void Test_Revisions_ContainerScopedOperationsAffectContentControlsAndTextBoxes() {
            string filePath = Path.Combine(_directoryWithFiles, "TrackedChangesContainerScopes.docx");
            CreateImportedReviewContainerDocument(filePath, "Scoped", "Alice");

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport contentControlReport = document.AcceptRevisions(new WordRevisionFilter {
                    RevisionType = WordReviewRevisionType.Insertion,
                    IsInContentControl = true
                });

                WordRevisionInfo contentControlRevision = Assert.Single(contentControlReport.MatchedRevisions);
                Assert.True(contentControlRevision.IsInContentControl);
                Assert.False(contentControlRevision.IsInTextBox);
                Assert.Equal("Scoped content-control insertion", contentControlRevision.AffectedText);

                WordRevisionOperationReport textBoxReport = document.RejectRevisions(new WordRevisionFilter {
                    RevisionType = WordReviewRevisionType.Insertion,
                    IsInTextBox = true
                });

                WordRevisionInfo textBoxRevision = Assert.Single(textBoxReport.MatchedRevisions);
                Assert.False(textBoxRevision.IsInContentControl);
                Assert.True(textBoxRevision.IsInTextBox);
                Assert.Equal("Scoped text-box insertion", textBoxRevision.AffectedText);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.DoesNotContain(review.Revisions, revision => revision.IsInContentControl);
                Assert.DoesNotContain(review.Revisions, revision => revision.IsInTextBox);
                Assert.Equal(2, review.CommentCount);
                Assert.Contains(review.Comments, comment => comment.IsInContentControl && comment.Text == "Scoped content-control comment.");
                Assert.Contains(review.Comments, comment => comment.IsInTextBox && comment.Text == "Scoped text-box comment.");
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                Body body = wordDocument.MainDocumentPart!.Document.Body!;

                Assert.Contains(body.Descendants<Run>(), run => run.InnerText == "Scoped content-control insertion");
                Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText == "Scoped content-control insertion");
                Assert.DoesNotContain(body.Descendants<Run>(), run => run.InnerText == "Scoped text-box insertion");
                Assert.DoesNotContain(body.Descendants<InsertedRun>(), run => run.InnerText == "Scoped text-box insertion");
            }
        }

        private static void CreateImportedFormattingRevisionDocument(string path, string label, string author) {
            File.Delete(path);
            DateTime revisionDate = new DateTime(2026, 6, 29, 13, 0, 0, DateTimeKind.Utc);

            using WordDocument document = WordDocument.Create(path);

            WordParagraph runParagraph = document.AddParagraph(label + " run formatting target");
            Run run = runParagraph._paragraph.Elements<Run>().First();
            run.RunProperties = new RunProperties(new Bold());
            run.RunProperties.RunPropertiesChange = new RunPropertiesChange(
                new PreviousRunProperties(new Italic())) {
                Id = "8201",
                Author = author,
                Date = revisionDate
            };

            WordParagraph paragraph = document.AddParagraph(label + " paragraph formatting target");
            paragraph._paragraph.ParagraphProperties = new ParagraphProperties(
                new Justification { Val = JustificationValues.Center });
            paragraph._paragraph.ParagraphProperties.ParagraphPropertiesChange = new ParagraphPropertiesChange(
                new PreviousParagraphProperties(new Justification { Val = JustificationValues.Right })) {
                Id = "8202",
                Author = author,
                Date = revisionDate
            };

            WordTable table = document.AddTable(1, 1);
            table.Rows[0].Cells[0].Paragraphs[0].SetText(label + " table formatting target");
            table.CheckTableProperties();
            table._tableProperties!.TableWidth = new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct };
            table._tableProperties.TablePropertiesChange = new TablePropertiesChange(
                new PreviousTableProperties(new TableWidth { Width = "2500", Type = TableWidthUnitValues.Pct })) {
                Id = "8203",
                Author = author,
                Date = revisionDate
            };

            WordTableRow row = table.Rows[0];
            row.AddTableRowProperties();
            row._tableRow.TableRowProperties!.Append(new TableRowHeight { Val = 720U, HeightType = HeightRuleValues.Exact });
            row._tableRow.TableRowProperties.Append(new TableRowPropertiesChange(
                new PreviousTableRowProperties(new TableRowHeight { Val = 240U, HeightType = HeightRuleValues.AtLeast })) {
                Id = "8204",
                Author = author,
                Date = revisionDate
            });

            WordTableCell cell = table.Rows[0].Cells[0];
            cell.AddTableCellProperties();
            cell._tableCellProperties!.RemoveAllChildren<TableCellWidth>();
            cell._tableCellProperties!.Append(new TableCellWidth { Width = "1800", Type = TableWidthUnitValues.Dxa });
            cell._tableCellProperties.Append(new TableCellPropertiesChange(
                new PreviousTableCellProperties(new TableCellWidth { Width = "1200", Type = TableWidthUnitValues.Dxa })) {
                Id = "8205",
                Author = author,
                Date = revisionDate
            });

            SectionProperties sectionProperties = document._document.Body!.Elements<SectionProperties>().FirstOrDefault()
                ?? document._document.Body.AppendChild(new SectionProperties());
            sectionProperties.RemoveAllChildren();
            sectionProperties.Append(
                new PageSize { Width = 12240U, Height = 15840U },
                new SectionPropertiesChange(
                    new PreviousSectionProperties(new PageSize { Width = 15840U, Height = 12240U })) {
                    Id = "8206",
                    Author = author,
                    Date = revisionDate
                });

            document.Save();
        }

        private static void AssertAcceptedFormatting(WordprocessingDocument wordDocument) {
            Body body = wordDocument.MainDocumentPart!.Document.Body!;

            Run run = Assert.Single(body.Descendants<Run>(), item => item.InnerText == "Accept run formatting target");
            Assert.NotNull(run.RunProperties!.Bold);
            Assert.Null(run.RunProperties.Italic);
            Assert.Null(run.RunProperties.RunPropertiesChange);

            Paragraph paragraph = Assert.Single(body.Descendants<Paragraph>(), item => item.InnerText == "Accept paragraph formatting target");
            Assert.Equal(JustificationValues.Center, paragraph.ParagraphProperties!.Justification!.Val!.Value);
            Assert.Null(paragraph.ParagraphProperties.ParagraphPropertiesChange);

            Table table = body.Descendants<Table>().Single();
            Assert.Equal("5000", table.TableProperties!.TableWidth!.Width!.Value);
            Assert.Null(table.TableProperties.TablePropertiesChange);

            TableRowProperties rowProperties = table.Descendants<TableRow>().First().TableRowProperties!;
            Assert.Equal(720U, rowProperties.GetFirstChild<TableRowHeight>()!.Val!.Value);
            Assert.Null(rowProperties.GetFirstChild<TableRowPropertiesChange>());

            TableCellProperties cellProperties = table.Descendants<TableCell>().First().TableCellProperties!;
            Assert.Equal("1800", cellProperties.GetFirstChild<TableCellWidth>()!.Width!.Value);
            Assert.Null(cellProperties.GetFirstChild<TableCellPropertiesChange>());

            SectionProperties sectionProperties = body.Elements<SectionProperties>().Single();
            Assert.Equal(12240U, sectionProperties.GetFirstChild<PageSize>()!.Width!.Value);
            Assert.Null(sectionProperties.GetFirstChild<SectionPropertiesChange>());
        }

        private static void AssertRejectedFormatting(WordprocessingDocument wordDocument) {
            Body body = wordDocument.MainDocumentPart!.Document.Body!;

            Run run = Assert.Single(body.Descendants<Run>(), item => item.InnerText == "Reject run formatting target");
            Assert.Null(run.RunProperties!.Bold);
            Assert.NotNull(run.RunProperties.Italic);
            Assert.Null(run.RunProperties.RunPropertiesChange);

            Paragraph paragraph = Assert.Single(body.Descendants<Paragraph>(), item => item.InnerText == "Reject paragraph formatting target");
            Assert.Equal(JustificationValues.Right, paragraph.ParagraphProperties!.Justification!.Val!.Value);
            Assert.Null(paragraph.ParagraphProperties.ParagraphPropertiesChange);

            Table table = body.Descendants<Table>().Single();
            Assert.Equal("2500", table.TableProperties!.TableWidth!.Width!.Value);
            Assert.Null(table.TableProperties.TablePropertiesChange);

            TableRowProperties rowProperties = table.Descendants<TableRow>().First().TableRowProperties!;
            TableRowHeight rowHeight = rowProperties.GetFirstChild<TableRowHeight>()!;
            Assert.Equal(240U, rowHeight.Val!.Value);
            Assert.Equal(HeightRuleValues.AtLeast, rowHeight.HeightType!.Value);
            Assert.Null(rowProperties.GetFirstChild<TableRowPropertiesChange>());

            TableCellProperties cellProperties = table.Descendants<TableCell>().First().TableCellProperties!;
            Assert.Equal("1200", cellProperties.GetFirstChild<TableCellWidth>()!.Width!.Value);
            Assert.Null(cellProperties.GetFirstChild<TableCellPropertiesChange>());

            SectionProperties sectionProperties = body.Elements<SectionProperties>().Single();
            Assert.Equal(15840U, sectionProperties.GetFirstChild<PageSize>()!.Width!.Value);
            Assert.Null(sectionProperties.GetFirstChild<SectionPropertiesChange>());
        }
    }
}

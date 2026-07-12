using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using W16Cid = DocumentFormat.OpenXml.Office2019.Word.Cid;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_InspectReview_ReadsCommentsRepliesResolvedStateAndTargets() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewInspection.Comments.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Commented target text").AddComment("Alice Reviewer", "AR", "Please review this.");
                WordComment parent = Assert.Single(document.Comments);
                parent.AddReply("Bob Reviewer", "BR", "Resolved in draft two.");
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                W15.CommentEx parentMetadata = mainPart.WordprocessingCommentsExPart!.CommentsEx!.Elements<W15.CommentEx>().First();
                parentMetadata.Done = true;
                mainPart.WordprocessingCommentsExPart.CommentsEx.Save();

                WordprocessingCommentsIdsPart commentsIdsPart = mainPart.AddNewPart<WordprocessingCommentsIdsPart>();
                commentsIdsPart.CommentsIds = new W16Cid.CommentsIds();
                commentsIdsPart.CommentsIds.Save();

                WordprocessingPeoplePart peoplePart = mainPart.AddNewPart<WordprocessingPeoplePart>();
                peoplePart.People = new W15.People();
                peoplePart.People.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.True(review.HasReviewMetadata);
                Assert.Equal(2, review.CommentCount);
                Assert.Equal(1, review.ReplyCount);
                Assert.Equal(1, review.ResolvedCommentCount);
                Assert.Contains(review.UnsupportedMetadata, detail => detail.Contains("commentsIds", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(review.UnsupportedMetadata, detail => detail.Contains("people", StringComparison.OrdinalIgnoreCase));

                WordCommentInfo parent = Assert.Single(review.Comments, comment => comment.Author == "Alice Reviewer");
                Assert.Equal("Please review this.", parent.Text);
                Assert.Equal("Commented target text", parent.TargetText);
                Assert.Equal(WordReviewLocationKind.Body, parent.TargetLocationKind);
                Assert.True(parent.IsResolved);
                Assert.False(parent.IsReply);

                WordCommentInfo reply = Assert.Single(review.Comments, comment => comment.Author == "Bob Reviewer");
                Assert.True(reply.IsReply);
                Assert.Equal(parent.ParaId, reply.ParentParaId);
                Assert.Null(reply.IsResolved);

                WordFeatureReport report = document.InspectFeatures();
                WordFeatureFinding comments = Assert.Single(report.FindFeatures("Comments"));
                Assert.Contains(comments.Details, detail => detail.Contains("Replies: 1", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(comments.Details, detail => detail.Contains("Resolved: 1", StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_InspectReview_ReadsRevisionsByTypeAuthorAndLocation() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewInspection.Revisions.docx");
            File.Delete(filePath);
            DateTime revisionDate = new DateTime(2026, 6, 28, 12, 0, 0, DateTimeKind.Utc);

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph paragraph = document.AddParagraph("Baseline ");
                paragraph.AddInsertedText("Added", "Alice", revisionDate);
                paragraph.AddDeletedText("Removed", "Bob", revisionDate);

                WordParagraph formatted = document.AddParagraph("Formatted paragraph");
                formatted._paragraph.ParagraphProperties ??= new ParagraphProperties();
                formatted._paragraph.ParagraphProperties.ParagraphPropertiesChange = new ParagraphPropertiesChange(
                    new ParagraphProperties(new Justification { Val = JustificationValues.Center })) {
                    Id = "77",
                    Author = "Carol",
                    Date = revisionDate
                };

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.Equal(3, review.RevisionCount);
                Assert.Single(review.GetRevisionsByAuthor("alice"));

                WordRevisionInfo insertion = Assert.Single(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.Insertion);
                Assert.Equal("Alice", insertion.Author);
                Assert.Equal("Added", insertion.AffectedText);
                Assert.Equal(WordReviewLocationKind.Body, insertion.LocationKind);

                WordRevisionInfo deletion = Assert.Single(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.Deletion);
                Assert.Equal("Bob", deletion.Author);
                Assert.Equal("Removed", deletion.AffectedText);

                WordRevisionInfo formatting = Assert.Single(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.ParagraphFormatting);
                Assert.Equal("Carol", formatting.Author);
                Assert.Contains("Formatted paragraph", formatting.LocationText, StringComparison.Ordinal);

                WordFeatureFinding revisions = Assert.Single(document.InspectFeatures().FindFeatures("Revisions"));
                Assert.Contains(revisions.Details, detail => detail.Contains("Insertion: 1", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(revisions.Details, detail => detail.Contains("Deletion: 1", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(revisions.Details, detail => detail.Contains("ParagraphFormatting: 1", StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_InspectReview_ReadsImportedReviewMetadataInsideContentControlsAndTextBoxes() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewInspection.ImportedContainers.docx");
            CreateImportedReviewContainerDocument(filePath, "Imported", "Alice Reviewer");

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.Equal(2, review.CommentCount);
                Assert.Equal(2, review.RevisionCount);

                WordCommentInfo contentControlComment = Assert.Single(review.Comments, comment => comment.Text == "Imported content-control comment.");
                Assert.Equal("Content-control target", contentControlComment.TargetText);
                Assert.Equal(WordReviewLocationKind.Body, contentControlComment.TargetLocationKind);
                Assert.True(contentControlComment.IsInContentControl);
                Assert.False(contentControlComment.IsInTextBox);

                WordCommentInfo textBoxComment = Assert.Single(review.Comments, comment => comment.Text == "Imported text-box comment.");
                Assert.Equal("Text-box target", textBoxComment.TargetText);
                Assert.Equal(WordReviewLocationKind.Body, textBoxComment.TargetLocationKind);
                Assert.False(textBoxComment.IsInContentControl);
                Assert.True(textBoxComment.IsInTextBox);

                WordRevisionInfo contentControlRevision = Assert.Single(review.Revisions, revision => revision.AffectedText == "Imported content-control insertion");
                Assert.Equal(WordReviewRevisionType.Insertion, contentControlRevision.RevisionType);
                Assert.True(contentControlRevision.IsInContentControl);
                Assert.False(contentControlRevision.IsInTextBox);

                WordRevisionInfo textBoxRevision = Assert.Single(review.Revisions, revision => revision.AffectedText == "Imported text-box insertion");
                Assert.Equal(WordReviewRevisionType.Insertion, textBoxRevision.RevisionType);
                Assert.False(textBoxRevision.IsInContentControl);
                Assert.True(textBoxRevision.IsInTextBox);

                WordReviewReport report = document.InspectReviewReport();
                string json = report.ToJson();
                Assert.Contains("\"isInContentControl\": true", json, StringComparison.Ordinal);
                Assert.Contains("\"isInTextBox\": true", json, StringComparison.Ordinal);

                string markdown = report.ToMarkdown();
                Assert.Contains("content-control", markdown, StringComparison.Ordinal);
                Assert.Contains("text-box", markdown, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_InspectReview_ReadsImportedMoveAndFormattingRevisions() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewInspection.ImportedMoveFormatting.docx");
            CreateImportedMoveAndFormattingRevisionDocument(filePath, "Imported", "Alice Reviewer");

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.Equal(5, review.RevisionCount);

                WordRevisionInfo moveFrom = Assert.Single(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.MoveFrom);
                Assert.Equal("Alice Reviewer", moveFrom.Author);
                Assert.Equal("Imported moved from", moveFrom.AffectedText);
                Assert.Equal(WordReviewLocationKind.Body, moveFrom.LocationKind);

                WordRevisionInfo moveTo = Assert.Single(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.MoveTo);
                Assert.Equal("Alice Reviewer", moveTo.Author);
                Assert.Equal("Imported moved to", moveTo.AffectedText);

                WordRevisionInfo runFormatting = Assert.Single(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.RunFormatting);
                Assert.Equal("Alice Reviewer", runFormatting.Author);
                Assert.Contains("Run formatting target", runFormatting.LocationText, StringComparison.Ordinal);
                Assert.False(runFormatting.IsInTable);

                WordRevisionInfo tableFormatting = Assert.Single(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.TableFormatting);
                Assert.True(tableFormatting.IsInTable);

                WordRevisionInfo cellFormatting = Assert.Single(review.Revisions, revision => revision.RevisionType == WordReviewRevisionType.TableCellFormatting);
                Assert.True(cellFormatting.IsInTable);
                Assert.Contains("Imported cell formatting target", cellFormatting.LocationText, StringComparison.Ordinal);

                WordReviewReport report = document.InspectReviewReport();
                string json = report.ToJson();
                Assert.Contains("\"revisionType\": \"MoveFrom\"", json, StringComparison.Ordinal);
                Assert.Contains("\"revisionType\": \"MoveTo\"", json, StringComparison.Ordinal);
                Assert.Contains("\"revisionType\": \"RunFormatting\"", json, StringComparison.Ordinal);
                Assert.Contains("\"revisionType\": \"TableFormatting\"", json, StringComparison.Ordinal);
                Assert.Contains("\"revisionType\": \"TableCellFormatting\"", json, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_InspectReview_ReadsImportedReviewMetadataInHeadersFootersAndNotes() {
            string filePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "imported-related-part-comments-revisions.docx"));
            Assert.True(File.Exists(filePath), $"Missing imported related-part review fixture: {filePath}");

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.Equal(4, review.CommentCount);
                Assert.Equal(4, review.RevisionCount);

                AssertReviewCommentLocation(review, WordReviewLocationKind.Header, "Imported corpus header comment target", "Imported corpus header comment.");
                AssertReviewCommentLocation(review, WordReviewLocationKind.Footer, "Imported corpus footer comment target", "Imported corpus footer comment.");
                AssertReviewCommentLocation(review, WordReviewLocationKind.Footnote, "Imported corpus footnote comment target", "Imported corpus footnote comment.");
                AssertReviewCommentLocation(review, WordReviewLocationKind.Endnote, "Imported corpus endnote comment target", "Imported corpus endnote comment.");

                AssertReviewRevisionLocation(review, WordReviewLocationKind.Header, "Imported corpus header insertion");
                AssertReviewRevisionLocation(review, WordReviewLocationKind.Footer, "Imported corpus footer insertion");
                AssertReviewRevisionLocation(review, WordReviewLocationKind.Footnote, "Imported corpus footnote insertion");
                AssertReviewRevisionLocation(review, WordReviewLocationKind.Endnote, "Imported corpus endnote insertion");

                WordReviewReport report = document.InspectReviewReport();
                string json = report.ToJson();
                Assert.Contains("\"targetLocationKind\": \"Header\"", json, StringComparison.Ordinal);
                Assert.Contains("\"targetLocationKind\": \"Footer\"", json, StringComparison.Ordinal);
                Assert.Contains("\"targetLocationKind\": \"Footnote\"", json, StringComparison.Ordinal);
                Assert.Contains("\"targetLocationKind\": \"Endnote\"", json, StringComparison.Ordinal);

                string markdown = report.ToMarkdown();
                Assert.Contains("Header", markdown, StringComparison.Ordinal);
                Assert.Contains("Footer", markdown, StringComparison.Ordinal);
                Assert.Contains("Footnote", markdown, StringComparison.Ordinal);
                Assert.Contains("Endnote", markdown, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_InspectReview_ReportsExtensibleCommentMetadataAsUnsupported() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewInspection.CommentsExtensible.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Extensible comment target").AddComment("Alice Reviewer", "AR", "Classic comment remains readable.");
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                AddExtendedPart(mainPart,
                    "http://schemas.microsoft.com/office/2018/10/relationships/commentsExtensible",
                    "application/vnd.ms-word.commentsExtensible+xml",
                    "<w16cex:commentsExtensible xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" />");
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                WordCommentInfo comment = Assert.Single(review.Comments);
                Assert.Equal("Classic comment remains readable.", comment.Text);

                string detail = Assert.Single(review.UnsupportedMetadata, metadata =>
                    metadata.Contains("commentsExtensible", StringComparison.OrdinalIgnoreCase));
                Assert.Contains("preserved", detail, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("not yet parsed", detail, StringComparison.OrdinalIgnoreCase);

                WordReviewReport report = document.InspectReviewReport();
                Assert.Equal(1, report.UnsupportedMetadataCount);

                string json = report.ToJson();
                Assert.Contains("commentsExtensible", json, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("not yet parsed", json, StringComparison.OrdinalIgnoreCase);

                string markdown = report.ToMarkdown();
                Assert.Contains("commentsExtensible", markdown, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("not yet parsed", markdown, StringComparison.OrdinalIgnoreCase);
            }
        }

        [Fact]
        public void Test_InspectReview_ReadsWordAuthoredReviewFixture() {
            string filePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-comments-revisions.docx"));
            Assert.True(File.Exists(filePath), $"Missing Word-authored review fixture: {filePath}");

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.True(review.HasReviewMetadata);

                WordCommentInfo comment = Assert.Single(review.Comments);
                Assert.Equal("OfficeIMO Word Fixture", comment.Author);
                Assert.Equal("OWF", comment.Initials);
                Assert.Equal("Word-authored comment for review corpus.", comment.Text);
                Assert.Equal("Word-authored comment target", comment.TargetText);
                Assert.Equal(WordReviewLocationKind.Body, comment.TargetLocationKind);
                Assert.False(comment.IsReply);
                Assert.False(comment.IsResolved);
                Assert.False(string.IsNullOrWhiteSpace(comment.ParaId));

                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.Author == "OfficeIMO Word Fixture" &&
                    revision.AffectedText == "Word-authored deletion target");
                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.Author == "OfficeIMO Word Fixture" &&
                    revision.AffectedText == "Word-authored inserted revision");

                Assert.Contains(review.UnsupportedMetadata, metadata => metadata.Contains("commentsIds", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(review.UnsupportedMetadata, metadata => metadata.Contains("people", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(review.UnsupportedMetadata, metadata => metadata.Contains("commentsExtensible", StringComparison.OrdinalIgnoreCase));

                WordReviewReport report = document.InspectReviewReport();
                Assert.Equal(1, report.CommentCount);
                Assert.True(report.RevisionCount >= 2);
                Assert.True(report.UnsupportedMetadataCount >= 3);

                string json = report.ToJson();
                Assert.Contains("Word-authored comment for review corpus.", json, StringComparison.Ordinal);
                Assert.Contains("Word-authored inserted revision", json, StringComparison.Ordinal);
                Assert.Contains("commentsExtensible", json, StringComparison.OrdinalIgnoreCase);

                string markdown = report.ToMarkdown();
                Assert.Contains("Word-authored deletion target", markdown, StringComparison.Ordinal);
                Assert.Contains("Unsupported Review Metadata", markdown, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_InspectReview_ReadsWordComAuthoredBodyTableFixture() {
            string filePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-body-table-comments-revisions.docx"));
            Assert.True(File.Exists(filePath), $"Missing Word COM-authored review fixture: {filePath}");

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.True(review.HasReviewMetadata);
                Assert.Equal(2, review.CommentCount);
                Assert.True(review.RevisionCount >= 2);

                WordCommentInfo bodyComment = Assert.Single(review.Comments, comment =>
                    comment.Text == "Word COM body comment for review corpus.");
                Assert.Equal("OfficeIMO Word COM Fixture", bodyComment.Author);
                Assert.Equal("OWC", bodyComment.Initials);
                Assert.Equal("Word COM body comment target", bodyComment.TargetText);
                Assert.Equal(WordReviewLocationKind.Body, bodyComment.TargetLocationKind);
                Assert.False(bodyComment.IsInTable);

                WordCommentInfo tableComment = Assert.Single(review.Comments, comment =>
                    comment.Text == "Word COM table comment for review corpus.");
                Assert.Equal("Word COM table comment target", tableComment.TargetText);
                Assert.Equal(WordReviewLocationKind.Body, tableComment.TargetLocationKind);
                Assert.True(tableComment.IsInTable);

                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Insertion &&
                    revision.Author == "OfficeIMO Word COM Fixture" &&
                    revision.AffectedText == "Word COM inserted body revision");
                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.Deletion &&
                    revision.Author == "OfficeIMO Word COM Fixture" &&
                    revision.AffectedText == "Word COM deletion target");

                WordReviewReport report = document.InspectReviewReport();
                Assert.Equal(2, report.CommentCount);
                Assert.True(report.RevisionCount >= 2);

                string json = report.ToJson();
                Assert.Contains("Word COM table comment for review corpus.", json, StringComparison.Ordinal);
                Assert.Contains("\"isInTable\": true", json, StringComparison.Ordinal);

                string markdown = report.ToMarkdown();
                Assert.Contains("Word COM inserted body revision", markdown, StringComparison.Ordinal);
                Assert.Contains("Word COM deletion target", markdown, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_InspectReview_ReadsWordComAuthoredThreadedCommentFixture() {
            string filePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-threaded-comments.docx"));
            Assert.True(File.Exists(filePath), $"Missing Word COM-authored threaded comment fixture: {filePath}");

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, false)) {
                Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package));
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.True(review.HasReviewMetadata);
                Assert.Equal(3, review.CommentCount);

                WordCommentInfo parent = Assert.Single(review.Comments, comment =>
                    comment.Text == "Word COM parent comment for threaded corpus.");
                Assert.Equal("OfficeIMO Word COM Fixture", parent.Author);
                Assert.Equal("OWC", parent.Initials);
                Assert.Equal("Word COM threaded comment target", parent.TargetText);
                Assert.Equal(WordReviewLocationKind.Body, parent.TargetLocationKind);
                Assert.True(parent.IsResolved);
                Assert.False(string.IsNullOrWhiteSpace(parent.ParaId));
                Assert.True(string.IsNullOrWhiteSpace(parent.ParentParaId));

                WordCommentInfo reply = Assert.Single(review.Comments, comment =>
                    comment.Text == "Word COM reply comment for threaded corpus.");
                Assert.Equal(parent.ParaId, reply.ParentParaId);
                Assert.True(reply.IsResolved);

                WordCommentInfo unresolved = Assert.Single(review.Comments, comment =>
                    comment.Text == "Word COM unresolved comment for threaded corpus.");
                Assert.False(unresolved.IsResolved);
                Assert.Equal("Word COM unresolved comment target", unresolved.TargetText);

                WordReviewReport report = document.InspectReviewReport();
                WordCommentThreadInfo thread = Assert.Single(report.CommentThreads, item =>
                    item.Parent.Text == "Word COM parent comment for threaded corpus.");
                Assert.Single(thread.Replies);
                Assert.Single(report.UnresolvedThreads, item =>
                    item.Text == "Word COM unresolved comment for threaded corpus.");

                string json = report.ToJson();
                Assert.Contains("Word COM reply comment for threaded corpus.", json, StringComparison.Ordinal);
                Assert.Contains("\"isResolved\": true", json, StringComparison.Ordinal);
                Assert.Contains("\"isResolved\": false", json, StringComparison.Ordinal);

                string markdown = report.ToMarkdown();
                Assert.Contains("Word COM parent comment for threaded corpus.", markdown, StringComparison.Ordinal);
                Assert.Contains("Word COM unresolved comment for threaded corpus.", markdown, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_InspectReview_ReadsWordComAuthoredRelatedPartRevisionFixture() {
            string filePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-related-part-revisions.docx"));
            Assert.True(File.Exists(filePath), $"Missing Word COM-authored related-part review fixture: {filePath}");

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.True(review.HasReviewMetadata);
                Assert.Equal(0, review.CommentCount);
                Assert.True(review.RevisionCount >= 8);

                AssertReviewRevisionLocation(review, WordReviewLocationKind.Header, "Word COM header inserted revision");
                AssertReviewRevisionLocation(review, WordReviewLocationKind.Footer, "Word COM footer inserted revision");
                AssertReviewRevisionLocation(review, WordReviewLocationKind.Footnote, "Word COM footnote inserted revision");
                AssertReviewRevisionLocation(review, WordReviewLocationKind.Endnote, "Word COM endnote inserted revision");
                AssertReviewRevisionLocation(review, WordReviewLocationKind.Header, "Word COM header deletion target", WordReviewRevisionType.Deletion);
                AssertReviewRevisionLocation(review, WordReviewLocationKind.Footer, "Word COM footer deletion target", WordReviewRevisionType.Deletion);
                AssertReviewRevisionLocation(review, WordReviewLocationKind.Footnote, "Word COM footnote deletion target", WordReviewRevisionType.Deletion);
                AssertReviewRevisionLocation(review, WordReviewLocationKind.Endnote, "Word COM endnote deletion target", WordReviewRevisionType.Deletion);
                Assert.Contains(review.Revisions, revision => revision.Author == "OfficeIMO Word COM Fixture");

                WordReviewReport report = document.InspectReviewReport();
                Assert.True(report.RevisionCount >= 8);

                string json = report.ToJson();
                Assert.Contains("\"locationKind\": \"Header\"", json, StringComparison.Ordinal);
                Assert.Contains("\"locationKind\": \"Footer\"", json, StringComparison.Ordinal);
                Assert.Contains("\"locationKind\": \"Footnote\"", json, StringComparison.Ordinal);
                Assert.Contains("\"locationKind\": \"Endnote\"", json, StringComparison.Ordinal);

                string markdown = report.ToMarkdown();
                Assert.Contains("Word COM header inserted revision", markdown, StringComparison.Ordinal);
                Assert.Contains("Word COM endnote deletion target", markdown, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_InspectReview_ReadsWordAuthoredMoveAndFormattingFixture() {
            string filePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-move-formatting-revisions.docx"));
            Assert.True(File.Exists(filePath), $"Missing Word-authored move/formatting fixture: {filePath}");

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.True(review.HasReviewMetadata);

                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.MoveFrom &&
                    revision.Author == "OfficeIMO Word Fixture" &&
                    revision.AffectedText == "Word-authored move source");
                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.MoveTo &&
                    revision.Author == "OfficeIMO Word Fixture" &&
                    revision.AffectedText == "Word-authored move source");
                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.RunFormatting &&
                    revision.Author == "OfficeIMO Word Fixture" &&
                    revision.LocationText.Contains("Word-authored run formatting target", StringComparison.Ordinal));
                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.TableFormatting &&
                    revision.Author == "OfficeIMO Word Fixture" &&
                    revision.LocationText.Contains("Word-authored cell formatting target", StringComparison.Ordinal) &&
                    revision.IsInTable);
                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.TableCellFormatting &&
                    revision.Author == "OfficeIMO Word Fixture" &&
                    revision.LocationText.Contains("Word-authored cell formatting target", StringComparison.Ordinal) &&
                    revision.IsInTable);
                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.TableRowFormatting &&
                    revision.Author == "OfficeIMO Word Fixture" &&
                    revision.LocationText.Contains("Word-authored cell formatting target", StringComparison.Ordinal) &&
                    revision.IsInTable);
                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.TableFormatting &&
                    revision.ElementName == "tblGridChange");

                WordReviewReport report = document.InspectReviewReport();
                Assert.True(report.RevisionCount >= 7);

                string json = report.ToJson();
                Assert.Contains("\"revisionType\": \"MoveFrom\"", json, StringComparison.Ordinal);
                Assert.Contains("\"revisionType\": \"MoveTo\"", json, StringComparison.Ordinal);
                Assert.Contains("\"revisionType\": \"RunFormatting\"", json, StringComparison.Ordinal);
                Assert.Contains("\"revisionType\": \"TableFormatting\"", json, StringComparison.Ordinal);
                Assert.Contains("\"revisionType\": \"TableRowFormatting\"", json, StringComparison.Ordinal);
                Assert.Contains("\"revisionType\": \"TableCellFormatting\"", json, StringComparison.Ordinal);

                string markdown = report.ToMarkdown();
                Assert.Contains("Word-authored move source", markdown, StringComparison.Ordinal);
                Assert.Contains("Word-authored run formatting target", markdown, StringComparison.Ordinal);
                Assert.Contains("Word-authored cell formatting target", markdown, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_InspectReview_ReadsWordAuthoredParagraphAndSectionFormattingFixture() {
            string filePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-paragraph-section-formatting-revisions.docx"));
            Assert.True(File.Exists(filePath), $"Missing Word-authored paragraph/section formatting fixture: {filePath}");

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();

                Assert.True(review.HasReviewMetadata);

                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.ParagraphFormatting &&
                    revision.Author == "OfficeIMO Word Fixture" &&
                    revision.LocationText.Contains("Word-authored paragraph formatting target", StringComparison.Ordinal));
                Assert.Contains(review.Revisions, revision =>
                    revision.RevisionType == WordReviewRevisionType.SectionFormatting &&
                    revision.Author == "OfficeIMO Word Fixture" &&
                    revision.LocationText.Contains("Word-authored section formatting anchor", StringComparison.Ordinal));

                WordReviewReport report = document.InspectReviewReport();
                Assert.True(report.RevisionCount >= 2);

                string json = report.ToJson();
                Assert.Contains("\"revisionType\": \"ParagraphFormatting\"", json, StringComparison.Ordinal);
                Assert.Contains("\"revisionType\": \"SectionFormatting\"", json, StringComparison.Ordinal);

                string markdown = report.ToMarkdown();
                Assert.Contains("Word-authored paragraph formatting target", markdown, StringComparison.Ordinal);
                Assert.Contains("Word-authored section formatting anchor", markdown, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_InspectReview_MatchesImportedCommentExtensionMetadataByParagraphId() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewInspection.ImportedCommentsExOrder.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Parent target").AddComment("Alice Reviewer", "AR", "Parent comment.");
                WordComment parent = Assert.Single(document.Comments);
                parent.AddReply("Bob Reviewer", "BR", "Reply comment.");
                document.AddParagraph("Other target").AddComment("Carol Reviewer", "CR", "Other comment.");
                parent.MarkResolved();
                document.Save();
            }

            ReverseCommentsExOnly(filePath);

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordReviewInfo review = document.InspectReview();

                WordCommentInfo parent = Assert.Single(review.Comments, comment => comment.Author == "Alice Reviewer");
                Assert.True(parent.IsResolved);
                Assert.False(parent.IsReply);

                WordCommentInfo reply = Assert.Single(review.Comments, comment => comment.Author == "Bob Reviewer");
                Assert.True(reply.IsReply);
                Assert.Equal(parent.ParaId, reply.ParentParaId);

                WordCommentInfo other = Assert.Single(review.Comments, comment => comment.Author == "Carol Reviewer");
                Assert.False(other.IsResolved.GetValueOrDefault());
                Assert.False(other.IsReply);

                WordReviewReport report = document.InspectReviewReport();
                WordCommentThreadInfo thread = Assert.Single(report.CommentThreads, item => item.Parent.Author == "Alice Reviewer");
                Assert.Equal("Reply comment.", Assert.Single(thread.Replies).Text);
                Assert.Equal(2, report.CommentThreadCount);

                WordComment wrapperParent = Assert.Single(document.Comments, comment => comment.Author == "Alice Reviewer");
                Assert.True(wrapperParent.IsResolved);
                Assert.Equal("Reply comment.", Assert.Single(wrapperParent.Replies).Text);
                wrapperParent.MarkUnresolved();
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordCommentInfo parent = Assert.Single(document.InspectReview().Comments, comment => comment.Author == "Alice Reviewer");
                Assert.False(parent.IsResolved);
            }
        }

        [Fact]
        public void Test_WordComment_ResolvedStateOperationsRoundTripAndPreserveModernMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewOperations.ResolveComments.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Resolution target").AddComment("Alice Reviewer", "AR", "Resolve this thread.");
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                WordprocessingCommentsIdsPart commentsIdsPart = mainPart.AddNewPart<WordprocessingCommentsIdsPart>();
                commentsIdsPart.CommentsIds = new W16Cid.CommentsIds();
                commentsIdsPart.CommentsIds.Save();

                WordprocessingPeoplePart peoplePart = mainPart.AddNewPart<WordprocessingPeoplePart>();
                peoplePart.People = new W15.People();
                peoplePart.People.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordComment comment = Assert.Single(document.Comments);
                Assert.Null(comment.IsResolved);

                Assert.Same(comment, comment.MarkResolved());
                Assert.True(comment.IsResolved);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordReviewInfo review = document.InspectReview();
                WordCommentInfo comment = Assert.Single(review.Comments);

                Assert.True(comment.IsResolved);
                Assert.Contains(review.UnsupportedMetadata, detail => detail.Contains("commentsIds", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(review.UnsupportedMetadata, detail => detail.Contains("people", StringComparison.OrdinalIgnoreCase));

                document.Comments.Single().MarkUnresolved();
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordCommentInfo comment = Assert.Single(document.InspectReview().Comments);
                Assert.False(comment.IsResolved);
            }
        }

        [Fact]
        public void Test_WordComment_MarkResolvedAssignsGeneratedParaIdToLegacyCommentParagraph() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewOperations.ResolveLegacyComment.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Legacy target").AddComment("Alice Reviewer", "AR", "Legacy comment.");
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                Comment comment = Assert.Single(mainPart.WordprocessingCommentsPart!.Comments!.Elements<Comment>());
                comment.Elements<Paragraph>().Single().ParagraphId = null;
                mainPart.WordprocessingCommentsPart.Comments.Save();

                W15.CommentsEx commentsEx = mainPart.WordprocessingCommentsExPart!.CommentsEx!;
                commentsEx.RemoveAllChildren<W15.CommentEx>();
                commentsEx.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordComment comment = Assert.Single(document.Comments);
                comment.MarkResolved();
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                Comment comment = Assert.Single(mainPart.WordprocessingCommentsPart!.Comments!.Elements<Comment>());
                string? paragraphId = comment.Elements<Paragraph>().Single().ParagraphId?.Value;
                Assert.NotNull(paragraphId);
                W15.CommentEx commentEx = Assert.Single(mainPart.WordprocessingCommentsExPart!.CommentsEx!.Elements<W15.CommentEx>());

                Assert.Equal(paragraphId, commentEx.ParaId?.Value);
                Assert.True(commentEx.Done?.Value);
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordCommentInfo comment = Assert.Single(document.InspectReview().Comments);
                Assert.True(comment.IsResolved);
                Assert.False(string.IsNullOrWhiteSpace(comment.ParaId));
            }
        }

        [Fact]
        public void Test_InspectReview_PreservesParagraphSeparatorsInCommentTargets() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewInspection.CommentRangeParagraphSeparators.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("First");
                document.AddParagraph("Second");
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                WordprocessingCommentsPart commentsPart = mainPart.WordprocessingCommentsPart ?? mainPart.AddNewPart<WordprocessingCommentsPart>();
                commentsPart.Comments = new Comments(
                    new Comment(new Paragraph(new Run(new Text("Range note")))) {
                        Id = "0",
                        Author = "Alice Reviewer",
                        Initials = "AR"
                    });
                commentsPart.Comments.Save();

                Paragraph[] paragraphs = mainPart.Document.Body!.Elements<Paragraph>().Take(2).ToArray();
                paragraphs[0].PrependChild(new CommentRangeStart { Id = "0" });
                paragraphs[1].Append(new CommentRangeEnd { Id = "0" });
                paragraphs[1].Append(new Run(new CommentReference { Id = "0" }));
                mainPart.Document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordCommentInfo comment = Assert.Single(document.InspectReview().Comments);

                Assert.Equal("First Second", comment.TargetText);
            }
        }

        [Fact]
        public void Test_WordComment_DeleteRemovesOnlyCommentReferenceFromSharedHeaderRun() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewOperations.DeleteHeaderCommentSharedRun.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.Header.Default!.AddParagraph("Header target").AddComment("Alice Reviewer", "AR", "Header note.");
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                Header header = Assert.Single(wordDocument.MainDocumentPart!.HeaderParts).Header;
                CommentReference reference = Assert.Single(header.Descendants<CommentReference>());
                Run run = Assert.IsType<Run>(reference.Parent);
                reference.Remove();
                run.RemoveAllChildren();
                run.Append(
                    new Text("Before ") { Space = SpaceProcessingModeValues.Preserve },
                    reference,
                    new Text(" after") { Space = SpaceProcessingModeValues.Preserve });
                header.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordComment comment = Assert.Single(document.Comments);
                comment.Remove();
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                Header header = Assert.Single(wordDocument.MainDocumentPart!.HeaderParts).Header;

                Assert.Empty(header.Descendants<CommentReference>());
                Assert.Contains(header.Descendants<Text>(), text => text.Text == "Before ");
                Assert.Contains(header.Descendants<Text>(), text => text.Text == " after");
            }
        }

        [Fact]
        public void Test_WordComment_DeleteThreadRemovesRepliesAndPreservesUnrelatedComments() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewOperations.DeleteThread.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Thread target").AddComment("Alice Reviewer", "AR", "Parent comment.");
                WordComment parent = Assert.Single(document.Comments);
                parent.AddReply("Bob Reviewer", "BR", "Reply comment.");
                document.AddParagraph("Unrelated target").AddComment("Carol Reviewer", "CR", "Keep this.");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordComment parent = Assert.Single(document.Comments, comment => comment.Author == "Alice Reviewer");
                Assert.Single(parent.Replies);

                parent.DeleteThread();
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                WordReviewInfo review = document.InspectReview();
                WordCommentInfo remaining = Assert.Single(review.Comments);

                Assert.Equal("Carol Reviewer", remaining.Author);
                Assert.Equal("Keep this.", remaining.Text);
                Assert.Equal("Unrelated target", remaining.TargetText);
                Assert.Equal(0, review.ReplyCount);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                Body body = wordDocument.MainDocumentPart!.Document.Body!;
                Assert.Single(body.Descendants<CommentReference>());
                Assert.Single(body.Descendants<CommentRangeStart>());
                Assert.Single(body.Descendants<CommentRangeEnd>());
            }
        }

        [Fact]
        public void Test_InspectReviewReport_ExportsJsonMarkdownAndActionDetails() {
            string filePath = Path.Combine(_directoryWithFiles, "ReviewReport.JsonMarkdown.docx");
            File.Delete(filePath);
            DateTime revisionDate = new DateTime(2026, 6, 28, 15, 0, 0, DateTimeKind.Utc);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Comment target").AddComment("Alice Reviewer", "AR", "Check | quote \"now\".");
                WordComment.GetAllComments(document).Single().AddReply("Bob Reviewer", "BR", "Reply with more detail.");

                WordParagraph paragraph = document.AddParagraph("Baseline ");
                paragraph.AddInsertedText("Accepted text", "Alice Reviewer", revisionDate);
                paragraph.AddDeletedText("Remaining deletion", "Bob Reviewer", revisionDate);
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                WordprocessingCommentsIdsPart commentsIdsPart = mainPart.AddNewPart<WordprocessingCommentsIdsPart>();
                commentsIdsPart.CommentsIds = new W16Cid.CommentsIds();
                commentsIdsPart.CommentsIds.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordRevisionOperationReport action = document.AcceptRevisions(new WordRevisionFilter {
                    Author = "Alice Reviewer",
                    RevisionType = WordReviewRevisionType.Insertion
                });

                WordReviewReport report = document.InspectReviewReport(action);

                Assert.Equal(2, report.CommentCount);
                Assert.Equal(1, report.CommentThreadCount);
                Assert.Equal(1, report.RevisionCount);
                Assert.Equal(1, report.UnresolvedThreadCount);
                Assert.Equal(1, report.ActionCount);
                Assert.True(report.UnsupportedMetadataCount >= 1);
                Assert.Contains(report.UnsupportedMetadata, detail => detail.Contains("commentsIds", StringComparison.OrdinalIgnoreCase));
                WordCommentThreadInfo thread = Assert.Single(report.CommentThreads);
                Assert.Equal("Check | quote \"now\".", thread.Parent.Text);
                Assert.Equal("Reply with more detail.", Assert.Single(thread.Replies).Text);

                using JsonDocument parsed = JsonDocument.Parse(report.ToJson());
                JsonElement root = parsed.RootElement;
                Assert.Equal(2, root.GetProperty("commentCount").GetInt32());
                Assert.Equal(1, root.GetProperty("commentThreadCount").GetInt32());
                Assert.Equal("Alice Reviewer", root.GetProperty("comments")[0].GetProperty("author").GetString());
                Assert.Equal("Check | quote \"now\".", root.GetProperty("comments")[0].GetProperty("text").GetString());
                Assert.Equal(1, root.GetProperty("commentThreads")[0].GetProperty("replyCount").GetInt32());
                Assert.Equal("Check | quote \"now\".", root.GetProperty("commentThreads")[0].GetProperty("parent").GetProperty("text").GetString());
                Assert.Equal("Bob Reviewer", root.GetProperty("commentThreads")[0].GetProperty("replies")[0].GetProperty("author").GetString());
                Assert.Equal("Deletion", root.GetProperty("revisions")[0].GetProperty("revisionType").GetString());
                Assert.Equal("Accept", root.GetProperty("actions")[0].GetProperty("operation").GetString());
                Assert.Equal("Accepted text", root.GetProperty("actions")[0].GetProperty("matchedRevisions")[0].GetProperty("affectedText").GetString());

                string markdown = report.ToMarkdown();
                Assert.Contains("# Word Review Report", markdown, StringComparison.Ordinal);
                Assert.Contains("## Unresolved Threads", markdown, StringComparison.Ordinal);
                Assert.Contains("## Comment Threads", markdown, StringComparison.Ordinal);
                Assert.Contains("Check \\| quote \"now\".", markdown, StringComparison.Ordinal);
                Assert.Contains("| 0 | Check \\| quote \"now\". | 1 | unknown |", markdown, StringComparison.Ordinal);
                Assert.Contains("| Accept | 1 |", markdown, StringComparison.Ordinal);
                Assert.Contains("commentsIds", markdown, StringComparison.OrdinalIgnoreCase);
            }
        }

        private static void CreateImportedReviewContainerDocument(string path, string label, string author) {
            File.Delete(path);
            DateTime revisionDate = new DateTime(2026, 6, 29, 10, 0, 0, DateTimeKind.Utc);

            using (WordDocument document = WordDocument.Create(path)) {
                WordParagraph contentControlParagraph = document.AddParagraph("Content-control target");
                contentControlParagraph.AddComment(author, "RV", label + " content-control comment.");
                contentControlParagraph.AddInsertedText(label + " content-control insertion", author, revisionDate);

                WordTextBox textBox = document.AddTextBox("Text-box target");
                WordParagraph textBoxParagraph = textBox.Paragraphs[0];
                textBoxParagraph.AddComment(author, "RV", label + " text-box comment.");
                textBoxParagraph.AddInsertedText(label + " text-box insertion", author, revisionDate);

                document.Save();
            }

            WrapFirstParagraphContentInRunContentControl(path);
        }

        private static void CreateImportedMoveAndFormattingRevisionDocument(string path, string label, string author) {
            File.Delete(path);
            DateTime revisionDate = new DateTime(2026, 6, 29, 11, 0, 0, DateTimeKind.Utc);

            using WordDocument document = WordDocument.Create(path);

            WordParagraph moveParagraph = document.AddParagraph("Move revisions ");
            moveParagraph._paragraph.Append(
                new MoveFromRun(new Run(new Text(label + " moved from") { Space = SpaceProcessingModeValues.Preserve })) {
                    Id = "8101",
                    Author = author,
                    Date = revisionDate
                },
                new MoveToRun(new Run(new Text(label + " moved to") { Space = SpaceProcessingModeValues.Preserve })) {
                    Id = "8102",
                    Author = author,
                    Date = revisionDate
                });

            WordParagraph runFormattingParagraph = document.AddParagraph("Run formatting target");
            Run run = runFormattingParagraph._paragraph.Elements<Run>().First();
            run.RunProperties = new RunProperties(new Bold());
            run.RunProperties.RunPropertiesChange = new RunPropertiesChange(
                new PreviousRunProperties(new Italic())) {
                Id = "8103",
                Author = author,
                Date = revisionDate
            };

            WordTable table = document.AddTable(1, 1);
            table.Rows[0].Cells[0].Paragraphs[0].SetText(label + " cell formatting target");
            table.CheckTableProperties();
            table._tableProperties!.TablePropertiesChange = new TablePropertiesChange(
                new PreviousTableProperties(new TableWidth { Width = "2500", Type = TableWidthUnitValues.Pct })) {
                Id = "8104",
                Author = author,
                Date = revisionDate
            };

            WordTableCell cell = table.Rows[0].Cells[0];
            cell.AddTableCellProperties();
            cell._tableCellProperties!.Append(new TableCellPropertiesChange(
                new PreviousTableCellProperties(new TableCellWidth { Width = "1200", Type = TableWidthUnitValues.Dxa })) {
                Id = "8105",
                Author = author,
                Date = revisionDate
            });

            document.Save();
        }

        private static void CreateImportedReviewRelatedPartDocument(string path, string label, string author) {
            File.Delete(path);
            DateTime revisionDate = new DateTime(2026, 6, 29, 12, 0, 0, DateTimeKind.Utc);

            using WordDocument document = WordDocument.Create(path);

            document.AddParagraph(label + " body anchor");
            document.AddHeadersAndFooters();

            WordParagraph headerParagraph = document.Header.Default!.AddParagraph(label + " header target");
            headerParagraph.AddComment(author, "RV", label + " header comment.");
            headerParagraph.AddInsertedText(label + " header insertion", author, revisionDate);

            WordParagraph footerParagraph = document.Footer.Default!.AddParagraph(label + " footer target");
            footerParagraph.AddComment(author, "RV", label + " footer comment.");
            footerParagraph.AddInsertedText(label + " footer insertion", author, revisionDate);

            document.AddParagraph(label + " footnote anchor").AddFootNote(label + " footnote target");
            WordParagraph footnoteParagraph = document.FootNotes.Last().Paragraphs!.Last();
            footnoteParagraph.AddComment(author, "RV", label + " footnote comment.");
            footnoteParagraph.AddInsertedText(label + " footnote insertion", author, revisionDate);

            document.AddParagraph(label + " endnote anchor").AddEndNote(label + " endnote target");
            WordParagraph endnoteParagraph = document.EndNotes.Last().Paragraphs!.Last();
            endnoteParagraph.AddComment(author, "RV", label + " endnote comment.");
            endnoteParagraph.AddInsertedText(label + " endnote insertion", author, revisionDate);

            document.Save();
        }

        private static void AssertReviewCommentLocation(WordReviewInfo review, WordReviewLocationKind locationKind, string targetText, string commentText) {
            WordCommentInfo comment = Assert.Single(review.Comments, item =>
                item.TargetLocationKind == locationKind &&
                item.Text == commentText);

            Assert.Contains(targetText, comment.TargetText, StringComparison.Ordinal);
            Assert.Contains(locationKind.ToString(), comment.TargetPartUri ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }

        private static void AssertReviewRevisionLocation(WordReviewInfo review, WordReviewLocationKind locationKind, string affectedText, WordReviewRevisionType revisionType = WordReviewRevisionType.Insertion) {
            WordRevisionInfo revision = Assert.Single(review.Revisions, item =>
                item.LocationKind == locationKind &&
                item.AffectedText == affectedText);

            Assert.Equal(revisionType, revision.RevisionType);
            Assert.Contains(locationKind.ToString(), revision.PartUri, StringComparison.OrdinalIgnoreCase);
        }

        private static void WrapFirstParagraphContentInRunContentControl(string path) {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            Paragraph paragraph = wordDocument.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
            OpenXmlElement[] children = paragraph.ChildElements
                .Select(child => child.CloneNode(true))
                .ToArray();

            paragraph.RemoveAllChildren();
            var content = new SdtContentRun();
            foreach (OpenXmlElement child in children) {
                content.Append(child);
            }

            paragraph.Append(new SdtRun(
                new SdtProperties(
                    new SdtAlias { Val = "Imported review content control" },
                    new Tag { Val = "ImportedReviewContentControl" },
                    new SdtId { Val = 9001 }),
                content));

            wordDocument.MainDocumentPart.Document.Save();
        }

        private static void ReverseCommentsExOnly(string path) {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
            W15.CommentsEx? commentsEx = mainPart.WordprocessingCommentsExPart?.CommentsEx;
            if (commentsEx == null) {
                return;
            }

            W15.CommentEx[] reversedCommentsEx = commentsEx.Elements<W15.CommentEx>()
                .Reverse()
                .Select(comment => (W15.CommentEx)comment.CloneNode(true))
                .ToArray();
            commentsEx.RemoveAllChildren<W15.CommentEx>();
            commentsEx.Append(reversedCommentsEx);
            commentsEx.Save();
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioCommentTests {
        [Fact]
        public void CommentsSaveLoadAndRoundTripAsNativePart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            DateTimeOffset created = new(2026, 1, 2, 3, 4, 5, TimeSpan.Zero);
            DateTimeOffset edited = new(2026, 1, 3, 4, 5, 6, TimeSpan.Zero);

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Review", 11, 8.5);
            VisioShape api = new("api", 3, 5, 2, 1, "API") {
                Name = "Process",
                NameU = "Process"
            };
            page.Shapes.Add(api);

            VisioComment comment = page.AddComment(api, "Review security controls", "Przemyslaw", "PK", new VisioCommentOptions {
                CreatedAt = created,
                EditedAt = edited,
                Done = true,
                AutoCommentType = 0
            });

            Assert.Equal("api", comment.ShapeId);
            Assert.Single(page.CommentsForShape(api));

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertNativeCommentPackage(filePath, "Review security controls", "Przemyslaw", "PK", "api", created, edited, done: true);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioComment loadedComment = Assert.Single(loaded.Pages.Single().Comments);
            Assert.Equal("Review security controls", loadedComment.Text);
            Assert.Equal("Przemyslaw", loadedComment.AuthorName);
            Assert.Equal("PK", loadedComment.AuthorInitials);
            Assert.Equal("api", loadedComment.ShapeId);
            Assert.True(loadedComment.Done);
            Assert.Equal(0, loadedComment.AutoCommentType);

            loaded.Save(roundTripPath);
            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertNativeCommentPackage(roundTripPath, "Review security controls", "Przemyslaw", "PK", "api", created, edited, done: true);
        }

        [Fact]
        public void FluentCanAddCommentsToLoadedPageById() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string editedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.AsFluent()
                .Page("Existing", page => page
                    .Rect("api", 3, 5, 2, 1, "API"))
                .End()
                .Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            loaded.AsFluent()
                .FirstPage(page => page
                    .CommentShape("api", "Needs owner", "Operations", "OP")
                    .Comment("Page review complete", "Operations", "OP"))
                .End()
                .Save(editedPath);

            VisioDocument roundTrip = VisioDocument.Load(editedPath);
            VisioPage pageAfterRoundTrip = Assert.Single(roundTrip.Pages);
            Assert.Equal(2, pageAfterRoundTrip.Comments.Count);
            Assert.Equal("Needs owner", Assert.Single(pageAfterRoundTrip.CommentsForShape("api")).Text);
            Assert.Contains(pageAfterRoundTrip.Comments, comment => comment.ShapeId == null && comment.Text == "Page review complete");
            AssertNativeCommentPackage(editedPath, "Needs owner", "Operations", "OP", "api", null, null, done: false);
        }

        [Fact]
        public void SaveAssignsFallbackIdsBackToManuallyAppendedComments() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Manual Comments", 11, 8.5);
            VisioComment manual = new("Manual list append") {
                AuthorName = "Operations",
                AuthorInitials = "OP"
            };
            page.Comments.Add(manual);

            document.Save();

            Assert.Equal(1, manual.Id);
            VisioComment addedAfterSave = page.AddComment("Added after save", "Operations", "OP");
            Assert.Equal(2, addedAfterSave.Id);
            document.Save();

            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument comments = ReadXml(archive, "visio/comments.xml");
            Assert.Equal(new[] { "1", "2" }, comments.Descendants(v + "CommentEntry")
                .Select(entry => (string?)entry.Attribute("IX"))
                .OrderBy(id => id)
                .ToArray());
        }

        [Fact]
        public void CommentsCanBeReviewedUpdatedReopenedAndRemovedInLoadedDocuments() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string reviewedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string fluentPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            DateTimeOffset created = new(2026, 2, 1, 8, 0, 0, TimeSpan.Zero);
            DateTimeOffset resolvedAt = new(2026, 2, 2, 9, 30, 0, TimeSpan.Zero);
            DateTimeOffset reopenedAt = new(2026, 2, 3, 10, 45, 0, TimeSpan.Zero);
            DateTimeOffset finalResolvedAt = new(2026, 2, 4, 11, 15, 0, TimeSpan.Zero);

            VisioDocument document = VisioDocument.Create(filePath);
            document.AsFluent()
                .Page("Review", page => page
                    .Rect("api", 3, 5, 2, 1, "API"))
                .End();
            VisioPage page = document.Pages.Single();
            VisioComment owner = page.AddCommentToShape("api", "Who owns this?", "Operations", "OP", new VisioCommentOptions { CreatedAt = created });
            VisioComment followUp = page.AddCommentToShape("api", "Check again later", "Operations", "OP", new VisioCommentOptions { CreatedAt = created });
            VisioComment temporary = page.AddComment("Remove after review", "Operations", "OP", new VisioCommentOptions { CreatedAt = created });
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages.Single();
            loadedPage.UpdateCommentText(owner.Id, "Owner confirmed", resolvedAt);
            loadedPage.ResolveComment(owner.Id, resolvedAt);
            loadedPage.ResolveComment(followUp.Id, resolvedAt);
            loadedPage.ReopenComment(followUp.Id, reopenedAt);
            Assert.True(loadedPage.RemoveComment(temporary.Id));
            loaded.Save(reviewedPath);

            VisioDocument reviewed = VisioDocument.Load(reviewedPath);
            VisioPage reviewedPage = reviewed.Pages.Single();
            Assert.Equal(2, reviewedPage.Comments.Count);
            Assert.Equal("Owner confirmed", reviewedPage.FindComment(owner.Id)!.Text);
            Assert.True(reviewedPage.FindComment(owner.Id)!.Done);
            Assert.False(reviewedPage.FindComment(followUp.Id)!.Done);
            Assert.Single(reviewedPage.ResolvedComments());
            Assert.Single(reviewedPage.UnresolvedComments());
            Assert.DoesNotContain(reviewedPage.Comments, comment => comment.Text == "Remove after review");
            AssertCommentXml(reviewedPath, "Owner confirmed", done: true, expectedEdited: resolvedAt);
            AssertCommentXml(reviewedPath, "Check again later", done: false, expectedEdited: reopenedAt);
            AssertCommentTextAbsent(reviewedPath, "Remove after review");

            reviewed.AsFluent()
                .FirstPage(pageBuilder => pageBuilder
                    .ShapeComments("api", comments => Assert.Equal(2, comments.Count))
                    .RemoveComment(owner.Id)
                    .UpdateComment(followUp.Id, "Follow-up resolved", finalResolvedAt)
                    .ResolveComment(followUp.Id, finalResolvedAt))
                .End()
                .Save(fluentPath);

            VisioDocument final = VisioDocument.Load(fluentPath);
            VisioComment remaining = Assert.Single(final.Pages.Single().Comments);
            Assert.Equal(followUp.Id, remaining.Id);
            Assert.Equal("Follow-up resolved", remaining.Text);
            Assert.True(remaining.Done);
            AssertCommentTextAbsent(fluentPath, "Owner confirmed");
            AssertCommentXml(fluentPath, "Follow-up resolved", done: true, expectedEdited: finalResolvedAt);
        }

        [Fact]
        public void LoadRejectsCommentsPartWithTooMuchXml() {
            string filePath = CreateDocumentWithComment();
            string oversizedCommentText = new('x', checked((int)VisioDocument.MaxCommentsXmlCharacters));
            ReplaceCommentsXml(filePath, CreateCommentsXml(oversizedCommentText));

            Assert.ThrowsAny<System.Xml.XmlException>(() => VisioDocument.Load(filePath));
        }

        [Fact]
        public void LoadRejectsTooManyNativeComments() {
            string filePath = CreateDocumentWithComment();
            ReplaceCommentsXml(filePath, CreateCommentsXml(Enumerable.Range(0, VisioDocument.MaxLoadedComments + 1)
                .Select(index => "Comment " + index.ToString())));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => VisioDocument.Load(filePath));
            Assert.Contains(VisioDocument.MaxLoadedComments.ToString(), exception.Message);
        }

        [Fact]
        public void LoadRejectsOversizedNativeCommentText() {
            string filePath = CreateDocumentWithComment();
            string oversizedCommentText = new('x', VisioDocument.MaxCommentTextCharacters + 1);
            ReplaceCommentsXml(filePath, CreateCommentsXml(oversizedCommentText));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => VisioDocument.Load(filePath));
            Assert.Contains(VisioDocument.MaxCommentTextCharacters.ToString(), exception.Message);
        }

        private static void AssertNativeCommentPackage(
            string filePath,
            string expectedText,
            string expectedAuthor,
            string expectedInitials,
            string expectedShapeId,
            DateTimeOffset? expectedCreated,
            DateTimeOffset? expectedEdited,
            bool done) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";

            XDocument contentTypes = ReadXml(archive, "[Content_Types].xml");
            Assert.Contains(contentTypes.Root!.Elements(ct + "Override"), element =>
                (string?)element.Attribute("PartName") == "/visio/comments.xml" &&
                (string?)element.Attribute("ContentType") == "application/vnd.ms-visio.comments+xml");

            XDocument rels = ReadXml(archive, "visio/_rels/document.xml.rels");
            Assert.Contains(rels.Root!.Elements(packageRelationships + "Relationship"), element =>
                (string?)element.Attribute("Type") == "http://schemas.microsoft.com/visio/2010/relationships/comments" &&
                (string?)element.Attribute("Target") == "comments.xml");

            XDocument page = ReadXml(archive, "visio/pages/page1.xml");
            string persistedShapeId = page.Descendants(v + "Shape")
                .Single(shape => shape.Element(v + "Text")?.Value == "API")
                .Attribute("ID")!
                .Value;

            XDocument comments = ReadXml(archive, "visio/comments.xml");
            XElement author = comments.Descendants(v + "AuthorEntry").Single();
            Assert.Equal(expectedAuthor, (string?)author.Attribute("Name"));
            Assert.Equal(expectedInitials, (string?)author.Attribute("Initials"));

            XElement entry = comments.Descendants(v + "CommentEntry")
                .Single(element => element.Value == expectedText);
            Assert.Equal("0", (string?)entry.Attribute("PageID"));
            Assert.Equal(persistedShapeId, (string?)entry.Attribute("ShapeID"));
            Assert.Equal(done ? "1" : "0", (string?)entry.Attribute("Done"));
            Assert.True(int.TryParse((string?)entry.Attribute("IX"), out _));
            Assert.True(int.TryParse((string?)entry.Attribute("AuthorID"), out _));

            if (expectedCreated.HasValue) {
                Assert.Equal(FormatExpectedDate(expectedCreated.Value), (string?)entry.Attribute("Date"));
            }

            if (expectedEdited.HasValue) {
                Assert.Equal(FormatExpectedDate(expectedEdited.Value), (string?)entry.Attribute("EditDate"));
            }

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Contains(loaded.Pages.Single().Comments, comment =>
                comment.Text == expectedText &&
                string.Equals(comment.ShapeId, expectedShapeId, StringComparison.Ordinal));
        }

        private static void AssertCommentXml(string filePath, string expectedText, bool done, DateTimeOffset expectedEdited) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument comments = ReadXml(archive, "visio/comments.xml");
            XElement entry = comments.Descendants(v + "CommentEntry")
                .Single(element => element.Value == expectedText);
            Assert.Equal(done ? "1" : "0", (string?)entry.Attribute("Done"));
            Assert.Equal(FormatExpectedDate(expectedEdited), (string?)entry.Attribute("EditDate"));
        }

        private static void AssertCommentTextAbsent(string filePath, string text) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument comments = ReadXml(archive, "visio/comments.xml");
            Assert.DoesNotContain(comments.Descendants(v + "CommentEntry"), entry => entry.Value == text);
        }

        private static string CreateDocumentWithComment() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Review", 11, 8.5);
            page.AddComment("Initial comment", "Operations", "OP");
            document.Save();
            return filePath;
        }

        private static void ReplaceCommentsXml(string filePath, string commentsXml) {
            using ZipArchive archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry entry = archive.GetEntry("visio/comments.xml") ?? throw new InvalidOperationException("Missing comments part.");
            entry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("visio/comments.xml", CompressionLevel.Optimal);
            using Stream stream = replacement.Open();
            using StreamWriter writer = new(stream, new UTF8Encoding(false));
            writer.Write(commentsXml);
        }

        private static string CreateCommentsXml(string commentText) {
            return CreateCommentsXml(new[] { commentText });
        }

        private static string CreateCommentsXml(IEnumerable<string> commentTexts) {
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            int id = 1;
            XDocument comments = new(new XElement(v + "Comments",
                new XAttribute("xmlns", v.NamespaceName),
                new XElement(v + "AuthorList",
                    new XElement(v + "AuthorEntry",
                        new XAttribute("ID", "1"),
                        new XAttribute("Name", "Operations"),
                        new XAttribute("Initials", "OP"))),
                new XElement(v + "CommentList",
                    commentTexts.Select(text => new XElement(v + "CommentEntry",
                        new XAttribute("IX", id++),
                        new XAttribute("AuthorID", "1"),
                        new XAttribute("PageID", "0"),
                        text)))));

            return comments.ToString(SaveOptions.DisableFormatting);
        }

        private static string FormatExpectedDate(DateTimeOffset value) {
            return System.Xml.XmlConvert.ToString(value.UtcDateTime, System.Xml.XmlDateTimeSerializationMode.Utc);
        }

        private static XDocument ReadXml(ZipArchive archive, string entryName) {
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}

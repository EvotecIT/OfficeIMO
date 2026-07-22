using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptCommentTests {
        [Fact]
        public void NativeWriter_AuthorsAndProjectsClassicComments() {
            DateTime firstDate = new(2026, 7, 15, 10, 11, 12, 345, DateTimeKind.Utc);
            DateTime secondDate = new(2025, 12, 24, 8, 9, 10, 120, DateTimeKind.Utc);
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide firstSlide = source.AddSlide();
                PowerPointSlide secondSlide = source.AddSlide();
                AddAuthors(source,
                    (0U, "Alice Reviewer", "AR", 3U),
                    (1U, "Bob Reviewer", "BR", 7U));
                AddComment(firstSlide, 0U, 3U, "Unicode review: żółć ✓",
                    firstDate, 4486, 1342);
                AddComment(secondSlide, 1U, 7U, "Second review",
                    secondDate, -25, 9021);

                LegacyPptWritePreflightReport report = source.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite, string.Join(Environment.NewLine, report.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptComment first = Assert.Single(legacy.Slides[0].Comments);
            Assert.Equal(3, first.Index);
            Assert.Equal("Alice Reviewer", first.Author);
            Assert.Equal("AR", first.Initials);
            Assert.Equal("Unicode review: żółć ✓", first.Text);
            Assert.Equal(firstDate, first.CreatedAtUtc);
            Assert.Equal(4486, first.X);
            Assert.Equal(1342, first.Y);
            LegacyPptComment second = Assert.Single(legacy.Slides[1].Comments);
            Assert.Equal(secondDate, second.CreatedAtUtc);
            Assert.Equal(-25, second.X);
            Assert.Equal(9021, second.Y);
            LegacyPptImportReport importReport = legacy.CreateImportReport();
            Assert.Equal(2, importReport.CommentCount);
            Assert.Equal(2, importReport.CommentAuthorCount);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            PowerPointReviewReport review = projected.InspectReviewComments();
            Assert.Equal(2, review.ClassicCount);
            Assert.Contains(review.Comments, comment => comment.AuthorName == "Alice Reviewer"
                && comment.Text == "Unicode review: żółć ✓" && comment.X == 4486
                && comment.Y == 1342 && comment.Created == firstDate);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void ImportedClassicCommentEdit_AppendsPreservingRecord() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide();
                AddAuthors(source, (0U, "Original Author", "OA", 1U));
                AddComment(slide, 0U, 1U, "Original review",
                    new DateTime(2026, 7, 1, 7, 0, 0, DateTimeKind.Utc), 120, 240);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] savedBytes;
            DateTime updatedDate = new(2026, 7, 15, 18, 30, 45, 678, DateTimeKind.Utc);
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                CommentAuthor author = imported.OpenXmlDocument.PresentationPart!
                    .CommentAuthorsPart!.CommentAuthorList!.Elements<CommentAuthor>().Single();
                author.Name = "Updated Author";
                author.Initials = "UA";
                author.LastIndex = 9U;
                Comment comment = imported.Slides[0].SlidePart.SlideCommentsPart!
                    .CommentList!.Elements<Comment>().Single();
                comment.Index = 9U;
                comment.DateTime = updatedDate;
                comment.Text!.Text = "Updated review with a longer body";
                comment.Position!.X = 777;
                comment.Position.Y = -333;

                LegacyPptWritePreflightReport report = imported.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite, string.Join(Environment.NewLine, report.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptComment commentResult = Assert.Single(Assert.Single(saved.Slides).Comments);
            Assert.Equal(9, commentResult.Index);
            Assert.Equal("Updated Author", commentResult.Author);
            Assert.Equal("UA", commentResult.Initials);
            Assert.Equal("Updated review with a longer body", commentResult.Text);
            Assert.Equal(updatedDate, commentResult.CreatedAtUtc);
            Assert.Equal(777, commentResult.X);
            Assert.Equal(-333, commentResult.Y);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedClassicCommentAddAndRemove_AppendPreservingRecords() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide();
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] withComment;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                AddAuthors(imported, (0U, "Added Author", "AA", 4U));
                AddComment(imported.Slides[0], 0U, 4U, "Added after import",
                    new DateTime(2026, 7, 15, 12, 0, 0, DateTimeKind.Utc), 500, 600);
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                withComment = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation added = LegacyPptPresentation.Load(withComment);
            Assert.Equal("Added after import", Assert.Single(added.Slides[0].Comments).Text);
            Assert.True(added.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            byte[] withoutComment;
            using (var input = new MemoryStream(withComment, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PresentationPart presentationPart = imported.OpenXmlDocument.PresentationPart!;
                presentationPart.DeletePart(presentationPart.CommentAuthorsPart!);
                SlidePart slidePart = imported.Slides[0].SlidePart;
                slidePart.DeletePart(slidePart.SlideCommentsPart!);
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                withoutComment = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation removed = LegacyPptPresentation.Load(withoutComment);
            Assert.Empty(removed.Slides[0].Comments);
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    added.Package.DocumentStream.Length)
                .SequenceEqual(added.Package.DocumentStream));
        }

        [Fact]
        public void NativeWriter_BlocksModernThreadedComments() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            source.AddSlide();
            PresentationPart presentationPart = source.OpenXmlDocument.PresentationPart!;
            PowerPointAuthorsPart modernAuthors = presentationPart.AddNewPart<PowerPointAuthorsPart>();
            FeedXml(modernAuthors,
                "<p188:authorLst xmlns:p188=\"http://schemas.microsoft.com/office/powerpoint/2018/8/main\" />");

            LegacyPptWritePreflightReport report = source.AnalyzeLegacyPptWrite();

            LegacyPptWriteFinding finding = Assert.Single(report.Findings,
                item => item.Code == "PPT-WRITE-MODERN-COMMENTS");
            Assert.Contains("no native PowerPoint 97-2003 representation",
                finding.Description, StringComparison.Ordinal);
        }

        [Fact]
        public void NativeWriter_BlocksUnrepresentableClassicAuthorMetadata() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide();
            AddAuthors(source, (0U, "Author", "A", 1U));
            CommentAuthor author = source.OpenXmlDocument.PresentationPart!
                .CommentAuthorsPart!.CommentAuthorList!.Elements<CommentAuthor>().Single();
            author.ColorIndex = 4U;
            AddComment(slide, 0U, 1U, "Review",
                new DateTime(2026, 7, 15, 10, 0, 0, DateTimeKind.Utc), 10, 20);

            LegacyPptWritePreflightReport report = source.AnalyzeLegacyPptWrite();

            LegacyPptWriteFinding finding = Assert.Single(report.Findings,
                item => item.Code == "PPT-WRITE-COMMENTS");
            Assert.Contains("color index", finding.Description, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void LegacyEncryption_LoadEncryptedPropagatesAggregateImportLimits() {
            const string password = "encrypted-limits";
            byte[] encryptedBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide first = source.AddSlide();
                PowerPointSlide second = source.AddSlide();
                AddAuthors(source, (0U, "Reviewer", "R", 2U));
                AddComment(first, 0U, 1U, "First", DateTime.UtcNow,
                    10, 20);
                AddComment(second, 0U, 2U, "Second", DateTime.UtcNow,
                    20, 30);
                encryptedBytes = source.ToEncryptedBytes(password,
                    PowerPointFileFormat.Ppt);
            }
            using var input = new MemoryStream(encryptedBytes,
                writable: false);

            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() =>
                PowerPointPresentation.LoadEncrypted(input, password,
                    new PowerPointLoadOptions {
                        LegacyPptImportOptions =
                            new LegacyPptImportOptions {
                                MaxCommentCount = 1
                            }
                    }));

            Assert.Contains("comment count", exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        private static void AddAuthors(PowerPointPresentation presentation,
            params (uint Id, string Name, string Initials, uint LastIndex)[] authors) {
            CommentAuthorsPart part = presentation.OpenXmlDocument.PresentationPart!
                .AddNewPart<CommentAuthorsPart>();
            part.CommentAuthorList = new CommentAuthorList(authors.Select(author =>
                new CommentAuthor {
                    Id = author.Id,
                    Name = author.Name,
                    Initials = author.Initials,
                    LastIndex = author.LastIndex,
                    ColorIndex = author.Id
                }));
        }

        private static void AddComment(PowerPointSlide slide, uint authorId, uint index,
            string text, DateTime created, long x, long y) {
            SlideCommentsPart part = slide.SlidePart.AddNewPart<SlideCommentsPart>();
            part.CommentList = new CommentList(new Comment(
                new Position { X = x, Y = y }, new Text(text)) {
                AuthorId = authorId,
                Index = index,
                DateTime = created
            });
        }

        private static void FeedXml(OpenXmlPart part, string xml) {
            using var data = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            part.FeedData(data);
        }
    }
}

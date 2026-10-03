using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void ThreadedComments_BatchPreservesRootReplyOrderPeopleAndSavedContent() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Review");
            string root = "30000000-0000-0000-0000-000000000001";
            DateTime date = new(2026, 10, 1, 8, 0, 0, DateTimeKind.Utc);
            var results = sheet.AddThreadedComments(new[] {
                new ExcelThreadedCommentOptions { Address = "A1", Text = "Root ", Author = "Reviewer", Id = root, Date = date },
                new ExcelThreadedCommentOptions { Address = "A1", Text = "Reply", Author = "Owner", ParentId = root, Date = date.AddMinutes(1) },
                new ExcelThreadedCommentOptions { Address = "B2", Text = "Other root", Author = "Reviewer", Date = date.AddMinutes(2), Done = true }
            });
            Assert.Equal(3, results.Count);
            Assert.Equal(results[0].PersonId, results[2].PersonId);
            Assert.NotEqual(results[0].PersonId, results[1].PersonId);
            Assert.False(results[0].IsReply); Assert.True(results[1].IsReply); Assert.True(results[2].Done);
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            var comments = reopened["Review"].GetThreadedComments();
            Assert.Equal(3, comments.Count);
            Assert.Equal("Root ", comments.Single(c => c.Id == results[0].Id).Text);
            Assert.Equal(date, comments.Single(c => c.Id == results[0].Id).Date);
            Assert.Equal(results[0].Id, comments.Single(c => c.Id == results[1].Id).ParentId);
            Assert.Equal("Owner", comments.Single(c => c.Id == results[1].Id).Author);
        }

        [Fact]
        public void ThreadedComments_BatchRejectsLateDuplicateWithoutAppendingEarlierValidItems() {
            using var document = ExcelDocument.Create();
            ExcelSheet first = document.AddWorksheet("First"), second = document.AddWorksheet("Second");
            var existing = first.AddThreadedComment("A1", "Existing");
            Assert.Throws<InvalidOperationException>(() => second.AddThreadedComments(new[] {
                new ExcelThreadedCommentOptions { Address = "B2", Text = "Valid first" },
                new ExcelThreadedCommentOptions { Address = "C3", Text = "Duplicate existing", Id = existing.Id }
            }));
            Assert.Empty(second.GetThreadedComments());
            Assert.Empty(second.WorksheetPart.WorksheetThreadedCommentsParts);
            Assert.Single(first.GetThreadedComments());
        }

        [Fact]
        public void ThreadedComments_BatchRejectsReplyToAnotherCellBeforePackageMutation() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Review");
            string root = "30000000-0000-0000-0000-000000000001";
            Assert.Throws<ArgumentException>(() => sheet.AddThreadedComments(new[] {
                new ExcelThreadedCommentOptions { Address = "A1", Text = "Root", Id = root },
                new ExcelThreadedCommentOptions { Address = "B2", Text = "Invalid reply", ParentId = root }
            }));
            Assert.Empty(sheet.GetThreadedComments());
            Assert.Empty(sheet.WorksheetPart.WorksheetThreadedCommentsParts);
        }

        [Fact]
        public void ThreadedComments_LargeBatchKeepsDistinctIdsAndAuthorsAfterReopen() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Review");
            const int count = 10_000;
            sheet.AddThreadedComments(Enumerable.Range(1, count).Select(row => new ExcelThreadedCommentOptions {
                Address = "A" + row, Text = "Comment " + row, Author = "Reviewer " + row,
                Date = new DateTime(2026, 10, 1, 8, 0, 0, DateTimeKind.Utc)
            }));
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            var comments = reopened["Review"].GetThreadedComments();
            Assert.Equal(count, comments.Count);
            Assert.Equal(count, comments.Select(c => c.Id).Distinct().Count());
            Assert.Equal(count, comments.Select(c => c.PersonId).Distinct().Count());
            Assert.Equal("Reviewer 10000", comments.Single(c => c.CellReference == "A10000").Author);
            Assert.Equal("Comment 10000", comments.Single(c => c.CellReference == "A10000").Text);
        }
    }
}

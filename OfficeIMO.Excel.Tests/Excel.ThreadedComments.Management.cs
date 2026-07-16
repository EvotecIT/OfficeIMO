using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void ThreadedComments_CanBeReadUpdatedResolvedAndRemoved() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            const string rootId = "{10000000-0000-0000-0000-000000000001}";
            const string firstReplyId = "{10000000-0000-0000-0000-000000000002}";
            const string secondReplyId = "{10000000-0000-0000-0000-000000000003}";

            try {
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    ExcelSheet sheet = document.AddWorksheet("Review");
                    sheet.CellValue(2, 2, 1250d);
                    sheet.AddThreadedComment(new ExcelThreadedCommentOptions {
                        Address = "B2",
                        Text = "Please review.",
                        Author = "Reviewer",
                        Id = rootId,
                        Date = new DateTime(2026, 7, 16, 8, 0, 0, DateTimeKind.Utc)
                    });
                    sheet.ReplyToThreadedComment(
                        rootId,
                        "First reply.",
                        "Owner",
                        new DateTime(2026, 7, 16, 8, 5, 0, DateTimeKind.Utc),
                        firstReplyId);
                    sheet.ReplyToThreadedComment(rootId, "Second reply.", "Auditor", id: secondReplyId);

                    IReadOnlyList<ExcelThreadedCommentSnapshot> initial = sheet.GetThreadedComments();
                    Assert.Equal(3, initial.Count);
                    Assert.Equal("Reviewer", sheet.GetThreadedComment(rootId)!.Author);
                    Assert.Equal(rootId, sheet.GetThreadedComment(firstReplyId)!.ParentId);

                    ExcelThreadedCommentResult updated = sheet.UpdateThreadedComment(new ExcelThreadedCommentUpdateOptions {
                        Id = rootId,
                        Text = "Variance confirmed.",
                        Author = "Senior Reviewer",
                        Date = new DateTime(2026, 7, 16, 9, 0, 0, DateTimeKind.Utc),
                        Done = true
                    });
                    Assert.Equal("Senior Reviewer", updated.Author);
                    Assert.True(updated.Done);
                    Assert.True(sheet.GetThreadedComment(rootId)!.Done);
                    Assert.Equal("Variance confirmed.", sheet.GetThreadedComment(rootId)!.Text);

                    ExcelThreadedCommentResult reopened = sheet.SetThreadedCommentResolved(rootId, resolved: false);
                    Assert.False(reopened.Done);
                    document.Save();
                    Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                }

                using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                    ExcelSheet sheet = document["Review"];
                    ExcelThreadedCommentSnapshot persistedRoot = sheet.GetThreadedComment(rootId)!;
                    Assert.Equal("Variance confirmed.", persistedRoot.Text);
                    Assert.Equal("Senior Reviewer", persistedRoot.Author);
                    Assert.Equal(new DateTime(2026, 7, 16, 9, 0, 0, DateTimeKind.Utc), persistedRoot.Date);
                    Assert.False(persistedRoot.Done);

                    ExcelFeatureReport report = document.InspectFeatures();
                    ExcelFeatureFinding finding = Assert.Single(report.FindFeatures("Threaded comments"));
                    Assert.Equal(ExcelFeatureSupportLevel.PartiallyEditable, finding.SupportLevel);
                    Assert.Equal(3, finding.Count);
                    Assert.Same(report, report.EnsureNoAdvancedFeatures());

                    Assert.Throws<InvalidOperationException>(() => sheet.RemoveThreadedComment(rootId, removeReplies: false));
                    Assert.True(sheet.RemoveThreadedComment(firstReplyId));
                    Assert.Equal(2, sheet.RemoveThreadedCommentsAt("B2"));
                    Assert.Empty(sheet.GetThreadedComments());
                    Assert.Empty(sheet.WorksheetPart.WorksheetThreadedCommentsParts);
                    document.Save();
                }

                using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
                Assert.Empty(spreadsheet.WorkbookPart!.WorksheetParts.Single().WorksheetThreadedCommentsParts);
            } finally {
                TryDelete(filePath);
            }
        }

        [Fact]
        public void ThreadedComments_RejectDuplicateOrOrphanedConversationIds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            const string rootIdWithoutBraces = "20000000-0000-0000-0000-000000000001";
            const string rootId = "{20000000-0000-0000-0000-000000000001}";
            const string replyId = "{20000000-0000-0000-0000-000000000002}";
            const string missingId = "{20000000-0000-0000-0000-000000000099}";

            try {
                using ExcelDocument document = ExcelDocument.Create(filePath);
                ExcelSheet sheet = document.AddWorksheet("Review");
                ExcelThreadedCommentResult root = sheet.AddThreadedComment(new ExcelThreadedCommentOptions {
                    Address = "C3",
                    Text = "Root",
                    Id = rootIdWithoutBraces
                });
                Assert.Equal(rootId, root.Id);

                Assert.Throws<InvalidOperationException>(() => sheet.AddThreadedComment(new ExcelThreadedCommentOptions {
                    Address = "C3",
                    Text = "Duplicate",
                    Id = rootId
                }));
                Assert.Throws<ArgumentException>(() => sheet.AddThreadedComment(new ExcelThreadedCommentOptions {
                    Address = "C3",
                    Text = "Orphan",
                    ParentId = missingId
                }));
                Assert.Throws<ArgumentException>(() => sheet.AddThreadedComment(new ExcelThreadedCommentOptions {
                    Address = "D3",
                    Text = "Wrong cell",
                    ParentId = rootId
                }));
                Assert.Throws<ArgumentException>(() => sheet.AddThreadedComment(new ExcelThreadedCommentOptions {
                    Address = "C3",
                    Text = "Invalid id",
                    Id = "not-a-guid"
                }));

                ExcelThreadedCommentResult reply = sheet.ReplyToThreadedComment(rootId, "Reply", id: replyId);
                Assert.True(reply.IsReply);
                Assert.Throws<ArgumentException>(() => sheet.ReplyToThreadedComment(replyId, "Nested reply"));
                Assert.Throws<KeyNotFoundException>(() => sheet.UpdateThreadedComment(new ExcelThreadedCommentUpdateOptions {
                    Id = missingId,
                    Text = "Missing"
                }));
            } finally {
                TryDelete(filePath);
            }
        }
    }
}

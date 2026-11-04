using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void DuplicateLegacyComments_DoNotCrashStructuralRowRemapping() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(3, 1).SetValue("Value");
            sheet.SetComment(3, 1, "Duplicate", author: "Tester");
            AppendDuplicateComment(sheet);

            sheet.InsertRows(2);

            string[] references = sheet.WorksheetPart.WorksheetCommentsPart!
                .Comments!.CommentList!.Elements<Comment>()
                .Select(comment => comment.Reference!.Value!)
                .ToArray();
            Assert.Equal(new[] { "A4", "A4" }, references);
            Assert.True(sheet.HasComment(4, 1));
        }

        [Fact]
        public void DuplicateLegacyComments_DoNotCrashSortedRowRemapping() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Name");
            sheet.Cell(1, 2).SetValue("Score");
            sheet.Cell(2, 1).SetValue("High");
            sheet.Cell(2, 2).SetValue(10D);
            sheet.Cell(3, 1).SetValue("Low");
            sheet.Cell(3, 2).SetValue(1D);
            sheet.SetComment(2, 1, "Duplicate", author: "Tester");
            AppendDuplicateComment(sheet);

            sheet.Range("A1:B3").SortByColumn(columnOffset: 2,
                ascending: true,
                hasHeader: true);

            string[] references = sheet.WorksheetPart.WorksheetCommentsPart!
                .Comments!.CommentList!.Elements<Comment>()
                .Select(comment => comment.Reference!.Value!)
                .ToArray();
            Assert.Equal(new[] { "A3", "A3" }, references);
            Assert.True(sheet.HasComment(3, 1));
        }

        private static void AppendDuplicateComment(ExcelSheet sheet) {
            Comment original = sheet.WorksheetPart.WorksheetCommentsPart!
                .Comments!.CommentList!.Elements<Comment>().Single();
            sheet.WorksheetPart.WorksheetCommentsPart.Comments.CommentList
                .Append((Comment)original.CloneNode(true));
            sheet.WorksheetPart.WorksheetCommentsPart.Comments.Save();
        }
    }
}

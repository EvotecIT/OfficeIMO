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
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(3, 1, "Value");
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
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(1, 2, "Score");
            sheet.CellValue(2, 1, "High");
            sheet.CellValue(2, 2, 10D);
            sheet.CellValue(3, 1, "Low");
            sheet.CellValue(3, 2, 1D);
            sheet.SetComment(2, 1, "Duplicate", author: "Tester");
            AppendDuplicateComment(sheet);

            sheet.SortRangeByColumn(
                "A1:B3",
                columnOffset: 2,
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

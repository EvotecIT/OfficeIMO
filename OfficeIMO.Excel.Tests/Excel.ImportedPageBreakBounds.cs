using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void ImportedUnsignedPageBreaksStayInPackageAndDoNotOverflowStructuralEdits() {
            string path = Path.Combine(_directoryWithFiles, "UnsignedPageBreaks.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                document.AddSheet("Data").Cell(1, 1).SetValue("Data");
                document.Save();
            }
            string unknownRowXml;
            string unknownColumnXml;
            using (var package = SpreadsheetDocument.Open(path, true)) {
                var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
                var rows = new RowBreaks { Count = 5U, ManualBreakCount = 4U };
                rows.Append(new Break { Id = 2U, Min = 0U, Max = 16383U, ManualPageBreak = true });
                rows.Append(new Break { Id = 1048575U, Min = 0U, Max = 16383U, ManualPageBreak = true });
                rows.Append(new Break { Id = 1048577U, Min = 0U, Max = 16383U, ManualPageBreak = true });
                var unknownRow = new Break { Id = uint.MaxValue, Min = 7U, Max = 12U, ManualPageBreak = true };
                rows.Append(unknownRow);
                rows.Append(new Break { Id = 7U, Min = 0U, Max = 16383U, ManualPageBreak = false });
                var columns = new ColumnBreaks { Count = 3U, ManualBreakCount = 3U };
                columns.Append(new Break { Id = 16384U, Min = 0U, Max = 1048575U, ManualPageBreak = true });
                columns.Append(new Break { Id = 16385U, Min = 0U, Max = 1048575U, ManualPageBreak = true });
                var unknownColumn = new Break { Id = uint.MaxValue, Min = 1U, Max = 8U, ManualPageBreak = true };
                columns.Append(unknownColumn);
                worksheet.Append(rows, columns);
                worksheet.Save();
                unknownRowXml = unknownRow.OuterXml;
                unknownColumnXml = unknownColumn.OuterXml;
            }

            using (var document = ExcelDocument.Load(path)) {
                var sheet = document.Sheets[0];
                Assert.Equal(new[] { 2, 1048575 }, sheet.GetManualRowPageBreaks());
                Assert.Equal(new[] { 16384 }, sheet.GetManualColumnPageBreaks());
                sheet.InsertRows(1);
                Assert.Equal(new[] { 3, 1048576 }, sheet.GetManualRowPageBreaks());
                Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));
                sheet.DeleteColumns(1);
                Assert.Equal(new[] { 16383 }, sheet.GetManualColumnPageBreaks());
                sheet.InsertColumns(1);
                Assert.Equal(new[] { 16384 }, sheet.GetManualColumnPageBreaks());
                Assert.Throws<InvalidOperationException>(() => sheet.InsertColumns(1));
                document.Save();
            }

            using var saved = SpreadsheetDocument.Open(path, false);
            var root = saved.WorkbookPart!.WorksheetParts.Single().Worksheet;
            var savedRows = root.GetFirstChild<RowBreaks>()!.Elements<Break>().ToArray();
            Assert.Equal(new uint[] { 3, 1048576, 1048577, uint.MaxValue, 8 }, savedRows.Select(item => item.Id!.Value));
            Assert.False(savedRows.Single(item => item.Id!.Value == 8U).ManualPageBreak!.Value);
            Assert.Equal(unknownRowXml, savedRows.Single(item => item.Id!.Value == uint.MaxValue).OuterXml);
            Assert.Equal(unknownColumnXml, root.GetFirstChild<ColumnBreaks>()!.Elements<Break>()
                .Single(item => item.Id!.Value == uint.MaxValue).OuterXml);
            Assert.Equal(new uint[] { 16384, 16385, uint.MaxValue }, root.GetFirstChild<ColumnBreaks>()!
                .Elements<Break>().Select(item => item.Id!.Value));
        }
    }
}

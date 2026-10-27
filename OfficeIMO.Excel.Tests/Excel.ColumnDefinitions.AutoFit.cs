using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void Test_GetColumnDefinitions_ReflectsAutoFitBeforeSave(bool streamBacked, bool selectedOnly) {
            using var source = new MemoryStream();
            using var output = new MemoryStream();
            using var document = streamBacked ? ExcelDocument.Create(source) : ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.CellValues(new[] {
                (1, 1, (object)"Description"), (1, 2, (object)"Amount"), (1, 3, (object)"Code")
            }, ExcelExecutionMode.Sequential);
            // Detached workbooks materialize the append and retain a nondeferred
            // direct-save candidate. Stream-backed workbooks can defer the append.
            sheet.CellValues(new[] {
                (2, 1, (object)"An extended service description that needs a wide column"),
                (2, 2, (object)119.88D), (2, 3, (object)"A"),
                (3, 1, (object)"Return credit"),
                (3, 2, (object)(-13.50D)), (3, 3, (object)"B")
            }, ExcelExecutionMode.Parallel);

            if (selectedOnly) sheet.AutoFitColumnsFor(new[] { 1 });
            else sheet.AutoFitColumns();

            ExcelColumnInfo[] beforeSave = sheet.GetColumnDefinitions().ToArray();
            Assert.Equal(selectedOnly ? new[] { 1 } : new[] { 1, 2, 3 },
                beforeSave.Select(column => column.StartIndex));
            Assert.All(beforeSave, column => {
                Assert.Equal(column.StartIndex, column.EndIndex);
                Assert.True(column.Width?.CharacterCount > 0D);
                Assert.True(column.CustomWidth);
            });
            Assert.True((beforeSave[0].Width?.CharacterCount) > 20D);
            if (!selectedOnly) Assert.True((beforeSave[0].Width?.CharacterCount) > (beforeSave[1].Width?.CharacterCount));

            ExcelColumnInfo[] inspected = Assert.Single(document.Inspect().Worksheets).Columns.ToArray();
            Assert.Equal(beforeSave.Select(column => (column.StartIndex, column.EndIndex, (column.Width?.CharacterCount), column.CustomWidth)),
                inspected.Select(column => (column.StartIndex, column.EndIndex, (column.Width?.CharacterCount), column.CustomWidth)));

            var workbookLock = document.EnsureLock();
            workbookLock.EnterReadLock();
            try {
                Assert.Equal(beforeSave.Select(column => (column.Width?.CharacterCount)), sheet.GetColumnDefinitions().Select(column => (column.Width?.CharacterCount)));
            } finally {
                workbookLock.ExitReadLock();
            }
            sheet.Batch(current => Assert.Equal(beforeSave.Select(column => (column.Width?.CharacterCount)),
                current.GetColumnDefinitions().Select(column => (column.Width?.CharacterCount))));

            document.Save(output);
            if (!streamBacked) Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            output.Position = 0;
            using var reopened = ExcelDocument.Load(output);
            ExcelColumnInfo[] afterSave = reopened.Sheets[0].GetColumnDefinitions().ToArray();
            Assert.Equal(beforeSave.Length, afterSave.Length);
            for (int i = 0; i < beforeSave.Length; i++) {
                Assert.Equal(beforeSave[i].StartIndex, afterSave[i].StartIndex);
                Assert.Equal(beforeSave[i].EndIndex, afterSave[i].EndIndex);
                // The direct package writer stores widths to three decimal places.
                Assert.InRange(System.Math.Abs((beforeSave[i].Width?.CharacterCount)!.Value - (afterSave[i].Width?.CharacterCount)!.Value), 0D, 0.001D);
                Assert.Equal(beforeSave[i].CustomWidth, afterSave[i].CustomWidth);
            }
            Assert.Equal("Return credit", reopened.Sheets[0].CellAt(3, 1).GetValue<string>());
            Assert.Equal(-13.50D, reopened.Sheets[0].CellAt(3, 2).GetValue<double>());
        }
    }
}

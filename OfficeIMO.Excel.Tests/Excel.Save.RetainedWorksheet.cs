using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(true, false)]
        [InlineData(false, false)]
        [InlineData(true, true)]
        [InlineData(false, true)]
        public void Save_retained_worksheet_handles_preserve_subsequent_values_styles_and_layout(bool sheetCaching, bool standardWriter) {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            document.SheetCachingEnabled = sheetCaching;
            ExcelSheet sheet = document.AddSheet("Report");
            sheet.Cell(1, 1).SetValue("Notes");
            sheet.Cell(2, 1).SetValue(123.45).SetNumberFormat("0.00");
            sheet.Cell(2, 1).SetBold();
            sheet.Cell(2, 1).SetHorizontalAlignment(ExcelHorizontalAlignment.Right);
            var options = new ExcelSaveOptions { DisableFastPackageWriter = standardWriter, ValidateOpenXml = true };
            document.Save(stream, options);

            // Keep the original handle and another alias alive across another package replacement.
            ExcelSheet alias = document.Sheets[0];
            sheet.Cell(2, 1).SetWrapText();
            alias.Cell(3, 1).SetValue("Added after save");
            alias.SetColumnWidth(1, ExcelColumnWidth.Characters(32));
            document.Save(stream, options);
            Assert.True(sheet.Cell(2, 1).GetStyle().WrapText);
            Assert.Equal("Added after save", sheet.Cell(3, 1).GetValue<string>());
            using (var saved = SpreadsheetDocument.Open(stream, false)) {
                Cell savedCell = saved.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference == "A2");
                CellFormat savedStyle = saved.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!
                    .Elements<CellFormat>().ElementAt((int)savedCell.StyleIndex!.Value);
                Assert.True(savedStyle.Alignment!.WrapText!.Value);
            }

            sheet.Cell(2, 1).SetWrapText(false);
            sheet.Cell(4, 1).SetValue(new DateTime(2026, 10, 16));
            document.Save(stream, options);
            using var package = SpreadsheetDocument.Open(stream, false);
            Worksheet worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
            Cell value = worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "A2");
            Assert.Equal("123.45", value.CellValue!.Text);
            CellFormat format = package.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!
                .Elements<CellFormat>().ElementAt((int)value.StyleIndex!.Value);
            Assert.False(format.Alignment!.WrapText?.Value ?? false);
            Assert.Equal(HorizontalAlignmentValues.Right, format.Alignment.Horizontal!.Value);
            Assert.NotEqual(0U, format.FontId!.Value);
            Assert.NotEqual(0U, format.NumberFormatId!.Value);
            Assert.Equal(32, worksheet.Descendants<Column>().Single(column => column.Min == 1U).Width!.Value);
            Cell date = worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "A4");
            Assert.Equal(new DateTime(2026, 10, 16).ToOADate(), double.Parse(date.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture));
        }

        [Theory]
        [InlineData(".xlsx")]
        [InlineData(".xltx")]
        public void Save_retained_file_worksheet_handle_preserves_edits_after_each_save(string extension) {
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-RetainedSheet-" + Guid.NewGuid().ToString("N") + extension);
            try {
                using (var document = ExcelDocument.Create(path)) {
                    ExcelSheet sheet = document.AddSheet("Report");
                    sheet.Cell(1, 1).SetValue("Before");
                    document.Save();
                    sheet.Cell(1, 1).SetValue("After");
                    sheet.Cell(1, 1).SetWrapText();
                    document.Save();
                    using var readback = ExcelDocument.Load(path);
                    Assert.Equal("After", readback.Sheets[0].Cell(1, 1).GetValue<string>());
                    Assert.True(readback.Sheets[0].Cell(1, 1).GetStyle().WrapText);
                }
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Save_does_not_rebind_a_removed_handle_after_transaction_rollback_and_sheet_id_reuse() {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            document.SheetCachingEnabled = false;
            document.AddSheet("Keep");
            ExcelSheet original = document.AddSheet("Original");
            original.Cell(1, 1).SetValue("Original");
            using var cancellation = new System.Threading.CancellationTokenSource();
            Assert.ThrowsAny<OperationCanceledException>(() => original.ApplyTransactionalMutation(token => {
                original.Cell(2, 1).SetValue("Interrupted");
                cancellation.Cancel();
                token.ThrowIfCancellationRequested();
                return 1;
            }, new ExcelMutationPlanOptions(), cancellation.Token));
            uint originalId = original.SheetElement.SheetId!.Value;
            document.RemoveWorksheet("Original");
            ExcelSheet replacement = document.AddSheet("Replacement");
            replacement.Cell(1, 1).SetValue("Replacement");
            Assert.Equal(originalId, replacement.SheetElement.SheetId!.Value);
            document.Save(stream);

            Assert.Throws<InvalidOperationException>(() => original.GetTables());
            Assert.NotSame(original.DeferredMetadataWorksheetPart, replacement.DeferredMetadataWorksheetPart);
            using var saved = ExcelDocument.Load(stream);
            Assert.Equal("Replacement", saved["Replacement"].Cell(1, 1).GetValue<string>());
        }
    }
}

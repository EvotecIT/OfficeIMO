using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("cell")]
        [InlineData("range")]
        [InlineData("newline")]
        [InlineData("richtext")]
        public void BulkMutation_wrapping_preserves_distinct_alignment_and_format_flags(string route) {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Alignment");
            sheet.CellValue(1, 1, "First");
            sheet.CellValue(2, 1, "Second");
            sheet.CellBold(1, 1);
            Stylesheet styles = document.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            Worksheet worksheet = document.OpenXmlDocument.WorkbookPart.WorksheetParts.Single().Worksheet;
            Cell[] cells = worksheet.Descendants<Cell>().ToArray();
            var expected = new List<CellFormat>();
            for (int index = 0; index < cells.Length; index++) {
                var format = (CellFormat)styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cells[0].StyleIndex!.Value).CloneNode(true);
                format.ApplyProtection = true;
                format.Protection = new Protection { Locked = false };
                format.Alignment = new Alignment {
                    Horizontal = index == 0 ? HorizontalAlignmentValues.Center : HorizontalAlignmentValues.Right,
                    Vertical = VerticalAlignmentValues.Top,
                    TextRotation = 25U,
                    Indent = 2U
                };
                cells[index].StyleIndex = (uint)styles.CellFormats.Count();
                styles.CellFormats.Append(format);
                styles.CellFormats.Count = (uint)styles.CellFormats.Count();
                var wrapped = (CellFormat)format.CloneNode(true);
                wrapped.Alignment!.WrapText = true;
                wrapped.ApplyAlignment = true;
                expected.Add(wrapped);
            }

            sheet.Batch(target => {
                if (route == "range") target.WrapCells(1, 2, 1);
                else for (int row = 1; row <= 2; row++) {
                    if (route == "newline") target.CellValue(row, 1, "First\nSecond");
                    else if (route == "richtext") target.CellAt(row, 1).SetRichText(new ExcelRichTextRun("First"), new ExcelRichTextRun("\nSecond"));
                    else target.CellWrapText(row, 1);
                }
            });
            document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });

            using var package = SpreadsheetDocument.Open(stream, false);
            CellFormat[] savedFormats = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ToArray();
            Cell[] savedCells = package.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToArray();
            for (int index = 0; index < savedCells.Length; index++) {
                Assert.Equal(expected[index].OuterXml, savedFormats[(int)savedCells[index].StyleIndex!.Value].OuterXml);
            }
        }

        [Fact]
        public void BulkMutation_nested_cross_sheet_styles_save_once_and_survive_a_callback_failure() {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            document.AddWorksheet("First").CellValue(1, 1, "Original");
            document.AddWorksheet("Second").CellValue(1, 1, "Other");
            document.Save(stream);
            ExcelSheet first = document.Sheets[0];
            ExcelSheet second = document.Sheets[1];
            first.CellBold(1, 1);
            Func<int> saves = TrackPartSaves(document, document.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!);
            Assert.Throws<InvalidOperationException>(() => first.Batch(target => {
                target.CellValue(1, 1, "Changed");
                second.Batch(other => {
                    other.CellBold(1, 1);
                    other.WrapCells(1, 1, 1);
                });
                throw new InvalidOperationException("Callback failed after editing");
            }));
            Assert.Equal(1, saves());
            document.Save(stream);
            using var reopened = ExcelDocument.Load(stream);
            Assert.Equal("Changed", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
            Assert.True(reopened.Sheets[1].CellAt(1, 1).GetStyle().Bold);
            Assert.True(reopened.Sheets[1].CellAt(1, 1).GetStyle().WrapText);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void BulkMutation_comments_tables_and_formulas_defer_parts_and_preserve_saved_content(bool sheetCaching) {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                document.SheetCachingEnabled = sheetCaching;
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "First");
                sheet.CellValue(2, 2, 7);
                WorksheetPart part = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single();
                Func<int> worksheetSaves = TrackPartSaves(document, part);
                sheet.AddTable("A1:B2", true, "ReportTable", ExcelTableStyle.TableStyleMedium2);
                TableDefinitionPart tablePart = part.TableDefinitionParts.Single();
                Func<int> tableSaves = TrackPartSaves(document, tablePart);
                sheet.SetTableStyle("ReportTable", ExcelTableStyle.TableStyleMedium4);
                sheet.SetTableStyle("ReportTable", ExcelTableStyle.TableStyleMedium9);
                sheet.CellFormula(3, 2, "SUM(B2:B2)");
                sheet.SetComment(2, 1, "First comment");
                WorksheetCommentsPart commentPart = part.WorksheetCommentsPart!;
                Func<int> commentSaves = TrackPartSaves(document, commentPart);
                sheet.SetComment(2, 1, "Replacement");
                sheet.SetComment(3, 1, "Removed");
                sheet.ClearComment(3, 1);
                Assert.Equal(0, worksheetSaves());
                Assert.Equal(0, tableSaves());
                Assert.Equal(0, commentSaves());
                document.Save(stream, new ExcelSaveOptions { DisableFastPackageWriter = true, ValidateOpenXml = true });
            }
            using (var package = SpreadsheetDocument.Open(stream, false)) {
                WorksheetPart part = package.WorkbookPart!.WorksheetParts.Single();
                Assert.Equal("TableStyleMedium9", part.TableDefinitionParts.Single().Table!.TableStyleInfo!.Name!.Value);
                Assert.Equal("SUM(B2:B2)", part.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "B3").CellFormula!.Text);
                Assert.Equal("Replacement", part.WorksheetCommentsPart!.Comments!.CommentList!.Elements<Comment>().Single().InnerText);
                string vmlId = part.Worksheet.GetFirstChild<LegacyDrawing>()!.Id!.Value!;
                using Stream vml = part.GetPartById(vmlId).GetStream();
                XDocument drawing = XDocument.Load(vml);
                XNamespace excel = "urn:schemas-microsoft-com:office:excel";
                Assert.Equal("1", drawing.Descendants(excel + "Row").Single().Value);
                Assert.Empty(new OpenXmlValidator().Validate(package));
            }
            using var reopened = ExcelDocument.Load(stream);
            Assert.Equal("Replacement", reopened.Sheets[0].GetComments().Single().Text);
            reopened.Sheets[0].ClearComment(2, 1);
            reopened.Sheets[0].SetComment(4, 1, "New part");
            reopened.Sheets[0].ClearComment(4, 1);
            reopened.Save(stream);
            using var cleared = SpreadsheetDocument.Open(stream, false);
            Assert.Null(cleared.WorkbookPart!.WorksheetParts.Single().WorksheetCommentsPart);
            Assert.Null(cleared.WorkbookPart.WorksheetParts.Single().Worksheet.GetFirstChild<LegacyDrawing>());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task BulkMutation_validated_fast_save_checks_encoded_output_before_destination_changes(bool asynchronous) {
            using var source = new MemoryStream();
            using var document = ExcelDocument.Create(source);
            OfficeIMO.Tests.ExcelPackageWriterFixtures.WithoutOptionalTheme(document);
            ExcelSheet sheet = document.AddWorksheet("Validated");
            sheet.CellValue(1, 1, "First");
            using var valid = new MemoryStream();
            var options = new ExcelSaveOptions { ValidateOpenXml = true };
            if (asynchronous) await document.SaveAsync(valid, options: options);
            else document.Save(valid, options);
            Assert.Equal(ExcelSavePackageWriter.SimplePackage, document.LastSaveDiagnostics.Writer);
            using (var package = SpreadsheetDocument.Open(valid, false)) Assert.Empty(new OpenXmlValidator().Validate(package));

            // A font cannot be a child of a cell format. Validate the fixture before testing destination safety.
            sheet.CellValue(2, 1, "Invalid");
            sheet.CellBold(2, 1);
            document.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().Last()
                .Append(new DocumentFormat.OpenXml.Spreadsheet.Font());
            Assert.NotEmpty(new OpenXmlValidator().Validate(document.OpenXmlDocument));
            byte[] original = { 1, 2, 3, 4, 5 };
            using var destination = new MemoryStream();
            destination.Write(original, 0, original.Length);
            destination.Position = 2;
            if (asynchronous) await Assert.ThrowsAsync<InvalidOperationException>(() => document.SaveAsync(destination, options: options));
            else Assert.Throws<InvalidOperationException>(() => document.Save(destination, options));
            Assert.Equal(original, destination.ToArray());
            Assert.Equal(2, destination.Position);
        }

        [Fact]
        public void BulkMutation_rollback_preserves_pending_comment_payloads_and_discards_failed_edits() {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Rollback");
            sheet.CellValue(1, 1, "Value");
            sheet.SetComment(1, 1, "Before transaction");
            Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.SetComment(1, 1, "Failed replacement");
                sheet.SetComment(2, 1, "Failed addition");
                throw new InvalidOperationException("Rollback");
            }, 2, new ExcelMutationPlanOptions(), default));
            document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            using var package = SpreadsheetDocument.Open(stream, false);
            WorksheetPart part = package.WorkbookPart!.WorksheetParts.Single();
            Assert.Equal("Before transaction", part.WorksheetCommentsPart!.Comments!.CommentList!.Elements<Comment>().Single().InnerText);
            using Stream vml = part.GetPartById(part.Worksheet.GetFirstChild<LegacyDrawing>()!.Id!.Value!).GetStream();
            XNamespace excel = "urn:schemas-microsoft-com:office:excel";
            Assert.Equal("0", XDocument.Load(vml).Descendants(excel + "Row").Single().Value);
        }

        [Theory]
        [InlineData("single")]
        [InlineData("batch")]
        [InlineData("converter")]
        public void BulkMutation_number_format_ids_follow_imported_ids_in_every_style_scope(string scope) {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Number formats");
            sheet.CellValue(1, 1, 12);
            sheet.CellBold(1, 1);
            Stylesheet styles = document.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            styles.NumberingFormats = new NumberingFormats(new NumberingFormat { NumberFormatId = 57U, FormatCode = "0.0" });
            void Format() => sheet.FormatCell(1, 1, "0.000");
            if (scope == "batch") sheet.Batch(_ => Format());
            else if (scope == "converter") { using (sheet.BeginStyleBatch()) Format(); }
            else Format();
            Assert.Equal(58U, styles.NumberingFormats.Elements<NumberingFormat>().Last().NumberFormatId!.Value);
            document.Save(stream);
            using var package = SpreadsheetDocument.Open(stream, false);
            Assert.Equal(58U, package.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.NumberingFormats!.Elements<NumberingFormat>().Last().NumberFormatId!.Value);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void BulkMutation_named_style_redefinition_invalidates_cached_base_formats_and_derivations(bool nested) {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Named style");
            ExcelSheet other = document.AddWorksheet("Other");
            sheet.CellBold(1, 1);
            sheet.CellItalic(1, 2);
            sheet.DefineNamedStyle("Input", 1, 1);
            sheet.ApplyNamedStyle("Input", "A2:A3");
            sheet.Batch(target => {
                target.CellWrapText(2, 1);
                if (nested) other.Batch(_ => target.DefineNamedStyle("Input", 1, 2));
                else target.DefineNamedStyle("Input", 1, 2);
                target.CellWrapText(3, 1);
            });
            document.Save(stream);
            using var reopened = ExcelDocument.Load(stream);
            ExcelCellStyleInfo applied = reopened.Sheets[0].CellAt(3, 1).GetStyle();
            Assert.True(applied.Italic);
            Assert.False(applied.Bold);
            Assert.True(applied.WrapText);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void BulkMutation_successive_imported_xlsb_formula_edits_mark_the_workbook_dirty(bool sheetCaching) {
            using var seedStream = new MemoryStream();
            using (var seed = ExcelDocument.Create(new MemoryStream())) {
                seed.AddWorksheet("Native").CellValue(1, 1, 1);
                seed.Save(seedStream, ExcelFileFormat.Xlsb);
            }
            using var document = ExcelDocument.Load(seedStream);
            document.SheetCachingEnabled = sheetCaching;
            using var first = new MemoryStream();
            document.Sheets.Single().CellFormula(1, 1, "SUM(1,2)");
            document.Save(first, ExcelFileFormat.Xlsb);
            using var second = new MemoryStream();
            document.Sheets.Single().CellFormula(1, 1, "SUM(3,4)");
            document.Save(second, ExcelFileFormat.Xlsb);
            using var reopened = ExcelDocument.Load(second);
            Assert.Equal("SUM(3,4)", reopened.Sheets.Single().GetFormulaText(1, 1));
        }

        [Fact]
        public void BulkMutation_native_xls_and_structural_planning_observe_pending_comment_drawings_inside_batch() {
            using var native = new MemoryStream();
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = document.AddWorksheet("Native comments");
                sheet.Batch(target => {
                    target.CellValue(1, 1, "Move with note");
                    target.SetComment(1, 1, "Pending note");
                    target.InsertRows(1);
                    document.Save(native, ExcelFileFormat.Xls);
                });
            }
            using var reopened = ExcelDocument.Load(native);
            Assert.True(reopened.Sheets.Single().HasComment(2, 1));
            Assert.Equal("Pending note", reopened.Sheets.Single().GetComments().Single().Text);
        }

        [Theory]
        [InlineData("name")]
        [InlineData("id")]
        public void BulkMutation_validated_save_rejects_missing_required_sheet_metadata_without_replacing_destination(string missing) {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddWorksheet("Invalid").CellValue(1, 1, "Value");
            Sheet sheet = document.OpenXmlDocument.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single();
            if (missing == "name") sheet.Name = null;
            else sheet.SheetId = null;
            using var destination = new MemoryStream();
            byte[] original = { 7, 8, 9 };
            destination.Write(original, 0, original.Length);
            Assert.Throws<InvalidOperationException>(() => document.Save(destination, new ExcelSaveOptions { ValidateOpenXml = true }));
            Assert.Equal(original, destination.ToArray());
        }
    }
}

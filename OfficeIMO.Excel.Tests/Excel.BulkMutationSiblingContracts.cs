using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void BulkMutation_table_filter_parts_defer_and_survive_wrapper_recreation(bool sheetCaching) {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Filters");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Number");
                sheet.CellValue(1, 3, "Flag");
                for (int row = 2; row <= 21; row++) {
                    sheet.CellValue(row, 1, "Row" + row);
                    sheet.CellValue(row, 2, row);
                }
                sheet.Range("A1:C21").CreateTable("Audit", hasHeader: true, ExcelTableStyle.TableStyleMedium2, includeAutoFilter: true);
                document.Save(stream);
                document.SheetCachingEnabled = sheetCaching;
                sheet = document.Sheets[0];
                WorksheetPart worksheetPart = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single();
                Func<int> tableSaves = TrackPartSaves(document, worksheetPart.TableDefinitionParts.Single());
                sheet.Batch(target => {
                    for (int count = 1; count <= 20; count++) {
                        target.AutoFilterBlanks("A1:C21", 2);
                        target.AutoFilterTopBottom("A1:C21", 1, (ushort)count, top: true, percent: false);
                    }
                });
                Assert.Equal(0, tableSaves());
                _ = document.Sheets[0];
                document.Save(stream);
                Assert.True(tableSaves() > 0);
            }
            stream.Position = 0;
            using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
            WorksheetPart saved = package.WorkbookPart!.WorksheetParts.Single();
            Assert.Null(saved.Worksheet.GetFirstChild<AutoFilter>());
            AutoFilter filter = saved.TableDefinitionParts.Single().Table!.GetFirstChild<AutoFilter>()!;
            Assert.Equal("A1:C21", filter.Reference!.Value);
            Assert.Equal(20D, filter.Descendants<Top10>().Single().Val!.Value);
            Assert.True(filter.Descendants<Filters>().Single().Blank!.Value);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void BulkMutation_hyperlink_routes_preserve_cell_and_font_formatting(bool batch) {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Links");
                for (int row = 1; row <= 3; row++) {
                    sheet.CellValue(row, 1, "Two\nlines");
                    sheet.CellAt(row, 1).SetBold(true).SetItalic(true).SetFontName("Arial").SetFontSize(14).SetNumberFormat("0.0000 \"units\"");
                    sheet.CellAlign(row, 1, ExcelHorizontalAlignment.Center);
                    sheet.CellVerticalAlign(row, 1, ExcelVerticalAlignment.Top);
                }
                WorkbookPart workbook = document.OpenXmlDocument.WorkbookPart!;
                Stylesheet styles = workbook.WorkbookStylesPart!.Stylesheet!;
                Cell[] cells = workbook.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToArray();
                foreach (Cell cell in cells) {
                    var format = (CellFormat)styles.CellFormats!.ElementAt((int)cell.StyleIndex!.Value).CloneNode(true);
                    format.Protection = new Protection { Locked = false, Hidden = true };
                    format.ApplyProtection = true;
                    format.QuotePrefix = true;
                    cell.StyleIndex = (uint)styles.CellFormats.ChildElements.Count;
                    styles.CellFormats.Append(format);
                    styles.CellFormats.Count = (uint)styles.CellFormats.ChildElements.Count;
                }
                void AddLinks(ExcelSheet target) {
                    target.SetHyperlink(1, 1, "https://example.org", "Two\nlines");
                    target.SetHyperlinkReference(2, 1, "https://example.org");
                    target.SetInternalLink(3, 1, "'Links'!B1", "Two\nlines");
                }
                if (batch) sheet.Batch(AddLinks); else AddLinks(sheet);
                document.Save(stream);
            }
            stream.Position = 0;
            using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
            WorkbookPart saved = package.WorkbookPart!;
            Stylesheet savedStyles = saved.WorkbookStylesPart!.Stylesheet!;
            foreach (Cell cell in saved.WorksheetParts.Single().Worksheet!.Descendants<Cell>()) {
                CellFormat format = (CellFormat)savedStyles.CellFormats!.ElementAt((int)cell.StyleIndex!.Value);
                Assert.True(format.Alignment!.WrapText!.Value);
                Assert.Equal(HorizontalAlignmentValues.Center, format.Alignment.Horizontal!.Value);
                Assert.Equal(VerticalAlignmentValues.Top, format.Alignment.Vertical!.Value);
                Assert.False(format.Protection!.Locked!.Value);
                Assert.True(format.Protection.Hidden!.Value);
                Assert.True(format.ApplyProtection!.Value);
                Assert.True(format.QuotePrefix!.Value);
                Assert.Contains(savedStyles.NumberingFormats!.Elements<NumberingFormat>(), item =>
                    item.NumberFormatId?.Value == format.NumberFormatId?.Value && item.FormatCode?.Value == "0.0000 \"units\"");
                var font = savedStyles.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)format.FontId!.Value);
                Assert.NotNull(font.Bold);
                Assert.NotNull(font.Italic);
                Assert.Equal("Arial", font.FontName!.Val!.Value);
                Assert.Equal(14D, font.FontSize!.Val!.Value);
                Assert.Equal("FF0563C1", font.Color!.Rgb!.Value);
                Assert.NotNull(font.Underline);
            }
        }

        [Fact]
        public void BulkMutation_gradient_named_and_conditional_styles_join_the_same_batch() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Styles");
                for (int row = 1; row <= 20; row++) sheet.CellValue(row, 1, row);
                sheet.CellBold(1, 1);
                WorkbookPart workbook = document.OpenXmlDocument.WorkbookPart!;
                Func<int> styleSaves = TrackPartSaves(document, workbook.WorkbookStylesPart!);
                Func<int> worksheetSaves = TrackPartSaves(document, workbook.WorksheetParts.Single());
                sheet.Batch(target => {
                    for (int row = 1; row <= 20; row++) {
                        target.CellGradientBackground(row, 1, "#FF0000", "#0000FF");
                        target.DefineNamedStyle("Row" + row, row, 1);
                        target.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                            Range = "A" + row, Type = "Expression", Formulas = new[] { "TRUE" },
                            DifferentialFillColorArgb = row % 2 == 0 ? "FF00FF00" : "FFFF0000"
                        });
                    }
                    document.DefineNamedStyle("Declared", new ExcelStyleDefinition { Bold = true, WrapText = true });
                    document.DefineNamedStyle("Removed", new ExcelStyleDefinition { Bold = true });
                    Assert.True(document.RemoveNamedStyle("Removed"));
                    target.ClearConditionalFormatting("A20");
                });
                Assert.Equal(1, styleSaves());
                Assert.Equal(0, worksheetSaves());
                Assert.Contains(document.GetNamedStyles(), style => style.Name == "Declared");
                Assert.DoesNotContain(document.GetNamedStyles(), style => style.Name == "Removed");
                document.Save(stream);
            }
            stream.Position = 0;
            using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
            WorkbookPart saved = package.WorkbookPart!;
            Assert.Equal(19, saved.WorksheetParts.Single().Worksheet!.Elements<ConditionalFormatting>().Count());
            Assert.Contains(saved.WorkbookStylesPart!.Stylesheet!.Fills!.Elements<Fill>(), fill => fill.GradientFill != null);
            Assert.Contains(saved.WorkbookStylesPart.Stylesheet.CellStyles!.Elements<CellStyle>(), style => style.Name?.Value == "Declared");
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void BulkMutation_merge_and_filter_edits_survive_batch_failure_and_structural_snapshot(bool failCallback) {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Layout");
                sheet.CellValue(1, 1, "Header");
                for (int row = 2; row <= 21; row++) sheet.CellValue(row, 1, "Value" + row);
                Func<int> worksheetSaves = TrackPartSaves(document, document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single());
                void Edit(ExcelSheet target) {
                    target.AutoFilterAdd("A1:C21");
                    for (int row = 2; row <= 21; row++) {
                        target.MergeRange($"A{row}:B{row}");
                        target.AutoFilterByHeaderEquals("Header", new[] { "Value" + row });
                    }
                    target.UnmergeRange("A21:B21");
                    target.MergeRange("A21:B21");
                    if (failCallback) throw new InvalidOperationException("Callback failure after edits");
                }
                if (failCallback) Assert.Throws<InvalidOperationException>(() => sheet.Batch(Edit));
                else sheet.Batch(Edit);
                Assert.Equal(0, worksheetSaves());
                sheet.InsertRows(2, 1);
                document.Save(stream);
            }
            stream.Position = 0;
            using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
            Worksheet worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            Assert.Equal(Enumerable.Range(3, 20).Select(row => $"A{row}:B{row}"),
                worksheet.GetFirstChild<MergeCells>()!.Elements<MergeCell>().Select(merge => merge.Reference!.Value));
            AutoFilter filter = worksheet.GetFirstChild<AutoFilter>()!;
            Assert.Equal("A1:C22", filter.Reference!.Value);
            Assert.Equal("Value21", filter.Descendants<Filter>().Single().Val!.Value);
        }
    }
}

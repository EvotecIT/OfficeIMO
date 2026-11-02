using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("date", false)]
        [InlineData("date", true)]
        [InlineData("offset", false)]
        [InlineData("offset", true)]
        [InlineData("duration", false)]
        [InlineData("duration", true)]
#if NET6_0_OR_GREATER
        [InlineData("date-only", false)]
        [InlineData("date-only", true)]
        [InlineData("time-only", false)]
        [InlineData("time-only", true)]
#endif
        public void BulkMutation_automatic_formats_follow_named_style_replacement_on_another_sheet(string kind, bool batched) {
            object value = AutomaticFormatValue(kind);
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                document.SheetCachingEnabled = false;
                ExcelSheet sheet = document.AddSheet("Values");
                ExcelSheet source = document.AddSheet("Styles");
                source.Cell(1, 1).SetBold();
                source.Cell(1, 2).SetItalic();
                source.DefineNamedStyle("Input", 1, 1);
                sheet.ApplyNamedStyle("Input", "A2:A3");
                void Edit() {
                    sheet.Cell(2, 1).SetValue(value);
                    if (batched) source.Batch(_ => source.DefineNamedStyle("Input", 1, 2));
                    else document.DefineNamedStyle("Input", new ExcelStyleDefinition { Italic = true });
                    sheet.Cell(3, 1).SetValue(value);
                }
                if (batched) sheet.Batch(_ => Edit());
                else Edit();
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }

            using (var reopened = ExcelDocument.Load(stream)) {
                ExcelCellStyleInfo applied = reopened.Sheets[0].Cell(3, 1).GetStyle();
                Assert.True(applied.Italic);
                Assert.False(applied.Bold);
            }
            using var saved = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = saved.WorkbookPart!;
            Sheet valuesSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name?.Value == "Values");
            Cell cell = ((WorksheetPart)workbookPart.GetPartById(valuesSheet.Id!.Value!)).Worksheet.Descendants<Cell>()
                .Single(item => item.CellReference?.Value == "A3");
            CellFormat format = saved.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!
                .Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
            Assert.Equal(kind == "duration" || kind == "time-only" ? 46U : 14U, format.NumberFormatId!.Value);
        }

        [Theory]
        [InlineData("date")]
        [InlineData("duration")]
        public void BulkMutation_automatic_format_cache_on_a_held_wrapper_survives_other_sheet_rollback(string kind) {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                document.SheetCachingEnabled = false;
                ExcelSheet held = document.AddSheet("Values");
                ExcelSheet other = document.AddSheet("Transaction");
                held.Cell(1, 1).SetValue("Stable");
                object value = AutomaticFormatValue(kind);
                Assert.Throws<InvalidOperationException>(() => other.ApplyTransactionalMutation(_ => {
                    held.Cell(1, 1).SetValue(value);
                    throw new InvalidOperationException("Rollback");
                }, 1, new ExcelMutationPlanOptions(), default));
                held.Cell(2, 1).SetValue(value);
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }

            using var saved = ExcelDocument.Load(stream);
            Assert.Equal("Stable", saved.Sheets[0].Cell(1, 1).GetValue<string>());
            Assert.NotEmpty(saved.Sheets[0].Cell(2, 1).GetFormattedText());
            using var package = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = package.WorkbookPart!;
            Sheet valuesSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name?.Value == "Values");
            Cell cell = ((WorksheetPart)workbookPart.GetPartById(valuesSheet.Id!.Value!)).Worksheet.Descendants<Cell>()
                .Single(item => item.CellReference?.Value == "A2");
            CellFormat format = package.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!
                .Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
            Assert.Equal(kind == "duration" ? 46U : 14U, format.NumberFormatId!.Value);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void BulkMutation_view_protection_and_replacement_defer_writes_and_preserve_saved_edits(bool failCallback) {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Settings");
                for (int row = 1; row <= 20; row++) sheet.Cell(row, 1).SetValue($"Old{row:D2}");
                Func<int> saves = TrackPartSaves(document, document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single());
                void Edit(ExcelSheet target) {
                    for (int row = 1; row <= 20; row++) {
                        target.SetViewOptions(zoomScale: (uint)(100 + row), rightToLeft: true);
                        target.Protect(new ExcelSheetProtectionOptions { AllowFormatCells = true });
                        Assert.Equal(1, target.ReplaceAll($"Old{row:D2}", $"New{row:D2}"));
                    }
                    if (failCallback) throw new InvalidOperationException("Callback failure");
                }
                if (failCallback) Assert.Throws<InvalidOperationException>(() => sheet.Batch(Edit));
                else sheet.Batch(Edit);
                Assert.Equal(0, saves());
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }

            using var reopened = ExcelDocument.Load(stream);
            ExcelSheet saved = reopened.Sheets[0];
            Assert.Equal(120U, saved.GetZoomScale());
            Assert.True(saved.RightToLeft);
            Assert.True(saved.IsProtected);
            for (int row = 1; row <= 20; row++) Assert.Equal($"New{row:D2}", saved.Cell(row, 1).GetValue<string>());
            using var package = SpreadsheetDocument.Open(stream, false);
            Assert.False(package.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetProtection>()!.FormatCells!.Value);
        }

        [Fact]
        public void BulkMutation_legacy_font_and_alignment_projection_share_the_stylesheet_batch() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Projection");
                sheet.Cell(1, 1).SetBold();
                Func<int> saves = TrackPartSaves(document, document.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!);
                sheet.Batch(target => {
                    for (int row = 1; row <= 20; row++) {
                        target.Cell(row, 1).SetValue("Value");
                        target.FormatCellFont(row, 1, "Calibri", 12, "FF112233", true, true, null);
                        target.FormatCellAlignment(row, 1, HorizontalAlignmentValues.Center, VerticalAlignmentValues.Center, true);
                    }
                });
                Assert.Equal(1, saves());
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }
            using var reopened = ExcelDocument.Load(stream);
            for (int row = 1; row <= 20; row++) {
                ExcelCellStyleInfo style = reopened.Sheets[0].Cell(row, 1).GetStyle();
                Assert.True(style.Bold);
                Assert.True(style.Italic);
                Assert.True(style.WrapText);
                Assert.Equal("center", style.HorizontalAlignment);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void BulkMutation_values_and_formulas_on_another_sheet_join_the_workbook_write_lock(bool materialized) {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet owner = document.AddSheet("Batch owner");
                ExcelSheet target = document.AddSheet("Values");
                if (materialized) {
                    target.Cell(1, 1).SetValue("Seed");
                    target.Cell(1, 1).SetBold();
                }
                owner.Batch(_ => {
                    target.Cell(1, 1).SetValue("Text");
                    target.Cell(2, 1).SetValue(2D);
                    target.Cell(3, 1).SetValue(3F);
                    target.Cell(4, 1).SetValue(4M);
                    target.Cell(5, 1).SetValue(5);
                    target.Cell(6, 1).SetValue(6L);
                    target.Cell(7, 1).SetValue((short)7);
                    target.Cell(8, 1).SetValue(8U);
                    target.Cell(9, 1).SetValue(9UL);
                    target.Cell(10, 1).SetValue((ushort)10);
                    target.Cell(11, 1).SetValue((byte)11);
                    target.Cell(12, 1).SetValue((sbyte)12);
                    target.Cell(13, 1).SetValue(true);
                    target.Cell(14, 1).SetValue(new DateTime(2026, 1, 2));
                    target.Cell(15, 1).SetValue(TimeSpan.FromHours(36));
                    target.Cell(16, 1).SetValue(new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.Zero));
                    target.CellValue<int>(17, 1, 17);
                    target.Cell(18, 1).SetValue((object)18D);
                    target.Cell(19, 1).SetValue(42D);
                    target.Cell(19, 1).SetFormula("6*7");
#if NET6_0_OR_GREATER
                    target.Cell(20, 1).SetValue(new DateOnly(2026, 1, 2));
                    target.Cell(21, 1).SetValue(new TimeOnly(3, 4, 5));
#endif
                });
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }
            using var reopened = ExcelDocument.Load(stream);
            ExcelSheet saved = reopened.Sheets.Single(sheet => sheet.Name == "Values");
            Assert.Equal("Text", saved.Cell(1, 1).GetValue<string>());
            for (int row = 2; row <= 12; row++) Assert.Equal((double)row, saved.Cell(row, 1).GetValue<double>());
            Assert.True(saved.Cell(13, 1).GetValue<bool>());
            for (int row = 14; row <= 16; row++) Assert.NotEmpty(saved.Cell(row, 1).GetFormattedText());
            Assert.Equal(17D, saved.Cell(17, 1).GetValue<double>());
            Assert.Equal(18D, saved.Cell(18, 1).GetValue<double>());
            Assert.Equal("6*7", saved.GetFormulaText(19, 1));
            Assert.Equal(42D, saved.Cell(19, 1).GetValue<double>());
#if NET6_0_OR_GREATER
            Assert.NotEmpty(saved.Cell(20, 1).GetFormattedText());
            Assert.NotEmpty(saved.Cell(21, 1).GetFormattedText());
#endif
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void BulkMutation_held_wrappers_share_header_and_search_cache_invalidation(bool replaceAll) {
            using var seed = new MemoryStream();
            using (var creator = ExcelDocument.Create(seed)) {
                ExcelSheet sheet = creator.AddSheet("Aliases");
                sheet.Cell(1, 1).SetValue("Old");
                sheet.Cell(1, 2).SetValue("Other");
                sheet.Cell(2, 1).SetValue("Value");
                creator.Save(seed);
            }
            using var document = ExcelDocument.Load(seed);
            document.SheetCachingEnabled = false;
            ExcelSheet held = document.Sheets.Single();
            Dictionary<string, int> originalSnapshot = held.GetHeaderMap();
            Assert.Equal(1, originalSnapshot["Old"]);
            Assert.Equal("A1", held.FindFirst("Old"));
            ExcelSheet alias = document.Sheets.Single();
            if (replaceAll) Assert.Equal(1, alias.ReplaceAll("Old", "New"));
            else alias.Cell(1, 1).SetValue("New");
            Dictionary<string, int> updated = held.GetHeaderMap();
            Assert.False(updated.ContainsKey("Old"));
            Assert.Equal(1, updated["New"]);
            Assert.Null(held.FindFirst("Old"));
            Assert.Equal("A1", held.FindFirst("New"));
            held.SetByHeader(2, "New", "Written");
            alias.Cell(1, 1).SetValue("Newest");
            Assert.Equal(1, held.GetHeaderMap()["Newest"]);
            Assert.True(originalSnapshot.ContainsKey("Old"));
            Assert.False(originalSnapshot.ContainsKey("Newest"));
            using var destination = new MemoryStream();
            document.Save(destination, new ExcelSaveOptions { ValidateOpenXml = true });
            using var reopened = ExcelDocument.Load(destination);
            Assert.Equal("Newest", reopened.Sheets[0].Cell(1, 1).GetValue<string>());
            Assert.Equal("Written", reopened.Sheets[0].Cell(2, 1).GetValue<string>());
        }

        [Fact]
        public void BulkMutation_held_wrapper_shared_string_indexes_bind_to_the_restored_table() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                document.SheetCachingEnabled = false;
                ExcelSheet held = document.AddSheet("Values");
                ExcelSheet owner = document.AddSheet("Transaction");
                held.Cell(1, 1).SetValue("Stable");
                Assert.Throws<InvalidOperationException>(() => owner.ApplyTransactionalMutation(_ => {
                    held.Cell(2, 1).SetValue("Transient");
                    throw new InvalidOperationException("Rollback");
                }, 1, new ExcelMutationPlanOptions(), default));
                owner.Cell(2, 1).SetValue("Replacement");
                held.Cell(2, 1).SetValue("Transient");
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }
            using var reopened = ExcelDocument.Load(stream);
            Assert.Equal("Transient", reopened.Sheets.Single(sheet => sheet.Name == "Values").Cell(2, 1).GetValue<string>());
            Assert.Equal("Replacement", reopened.Sheets.Single(sheet => sheet.Name == "Transaction").Cell(2, 1).GetValue<string>());
        }

        [Fact]
        public void BulkMutation_shared_header_cache_respects_date_options_converters_and_cancellation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Headers");
            sheet.Cell(1, 1).SetValue(45292D);
            sheet.Cell(2, 1).SetValue("Value");
            sheet.Cell(1, 1).SetNumberFormat("yyyy-mm-dd");
            var rawOptions = new ExcelReaderOptions { TreatDatesUsingNumberFormat = false };
            Assert.Equal(1, sheet.GetHeaderMap(rawOptions)["45292"]);
            Assert.False(sheet.GetHeaderMap().ContainsKey("45292"));
            Assert.Equal(1, sheet.GetHeaderMap(rawOptions)["45292"]);
            string converted = "Converted";
            var converterOptions = new ExcelReaderOptions {
                CellValueConverter = _ => new ExcelCellValue(converted)
            };
            Assert.Equal(1, sheet.GetHeaderMap(converterOptions)["Converted"]);
            converted = "Changed";
            Assert.Equal(1, sheet.GetHeaderMap(converterOptions)["Changed"]);
            Assert.False(sheet.GetHeaderMap().ContainsKey("Changed"));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => {
                sheet.GetHeaderMap(new ExcelReaderOptions { CancellationToken = cancellation.Token });
            });
        }

        [Theory]
        [InlineData("cell")]
        [InlineData("range")]
        [InlineData("apply")]
        [InlineData("redefine")]
        public void BulkMutation_header_cache_follows_number_format_and_named_style_changes(string route) {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Headers");
            ExcelSheet source = document.AddSheet("Styles");
            sheet.Cell(1, 1).SetValue(45292D);
            sheet.Cell(2, 1).SetValue("Value");
            source.Cell(1, 1).SetValue("Date style");
            source.Cell(1, 1).SetNumberFormat("yyyy-mm-dd");
            source.Cell(1, 2).SetValue("General style");
            source.DefineNamedStyle("Header", 1, route == "redefine" ? 2 : 1);
            if (route == "redefine") sheet.ApplyNamedStyle("Header", "A1:A1");
            Assert.Equal(1, sheet.GetHeaderMap()["45292"]);
            switch (route) {
                case "cell": sheet.Cell(1, 1).SetNumberFormat("yyyy-mm-dd"); break;
                case "range": sheet.Range("A1:A1").SetNumberFormat("yyyy-mm-dd"); break;
                case "apply": sheet.ApplyNamedStyle("Header", "A1:A1"); break;
                case "redefine": source.DefineNamedStyle("Header", 1, 1); break;
            }
            Dictionary<string, int> changed = sheet.GetHeaderMap();
            Assert.False(changed.ContainsKey("45292"));
            Assert.Single(changed);
            Assert.Equal(1, changed.Values.Single());
            sheet.RefreshHeaderCache();
            Assert.Equal(sheet.GetHeaderMap(), changed);
        }

        private static object AutomaticFormatValue(string kind) => kind switch {
            "date" => new DateTime(2026, 1, 2),
            "offset" => new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.Zero),
            "duration" => TimeSpan.FromHours(36),
#if NET6_0_OR_GREATER
            "date-only" => new DateOnly(2026, 1, 2),
            "time-only" => new TimeOnly(3, 4, 5),
#endif
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
    }
}

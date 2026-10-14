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
                ExcelSheet sheet = document.AddWorksheet("Values");
                ExcelSheet source = document.AddWorksheet("Styles");
                source.CellBold(1, 1);
                source.CellItalic(1, 2);
                source.DefineNamedStyle("Input", 1, 1);
                sheet.ApplyNamedStyle("Input", "A2:A3");
                void Edit() {
                    sheet.CellValue(2, 1, value);
                    if (batched) source.Batch(_ => source.DefineNamedStyle("Input", 1, 2));
                    else document.DefineNamedStyle("Input", new ExcelStyleDefinition { Italic = true });
                    sheet.CellValue(3, 1, value);
                }
                if (batched) sheet.Batch(_ => Edit());
                else Edit();
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }

            using (var reopened = ExcelDocument.Load(stream)) {
                ExcelCellStyleSnapshot applied = reopened.Sheets[0].CellAt(3, 1).GetStyle();
                Assert.True(applied.Italic);
                Assert.False(applied.Bold);
            }
            using var saved = SpreadsheetDocument.Open(stream, false);
            Cell cell = saved.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
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
                ExcelSheet held = document.AddWorksheet("Values");
                ExcelSheet other = document.AddWorksheet("Transaction");
                held.CellValue(1, 1, "Stable");
                object value = AutomaticFormatValue(kind);
                Assert.Throws<InvalidOperationException>(() => other.ApplyTransactionalMutation(_ => {
                    held.CellValue(1, 1, value);
                    throw new InvalidOperationException("Rollback");
                }, 1, new ExcelMutationPlanOptions(), default));
                held.CellValue(2, 1, value);
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }

            using var saved = ExcelDocument.Load(stream);
            Assert.Equal("Stable", saved.Sheets[0].CellAt(1, 1).GetValue<string>());
            Assert.NotEmpty(saved.Sheets[0].CellAt(2, 1).GetFormattedText());
            using var package = SpreadsheetDocument.Open(stream, false);
            Cell cell = package.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>()
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
                ExcelSheet sheet = document.AddWorksheet("Settings");
                for (int row = 1; row <= 20; row++) sheet.CellValue(row, 1, $"Old{row:D2}");
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
            for (int row = 1; row <= 20; row++) Assert.Equal($"New{row:D2}", saved.CellAt(row, 1).GetValue<string>());
            using var package = SpreadsheetDocument.Open(stream, false);
            Assert.False(package.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetProtection>()!.FormatCells!.Value);
        }

        [Fact]
        public void BulkMutation_legacy_font_and_alignment_projection_share_the_stylesheet_batch() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddWorksheet("Projection");
                sheet.CellBold(1, 1);
                Func<int> saves = TrackPartSaves(document, document.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!);
                sheet.Batch(target => {
                    for (int row = 1; row <= 20; row++) {
                        target.CellValue(row, 1, "Value");
                        target.FormatCellFont(row, 1, "Calibri", 12, "FF112233", true, true, null);
                        target.FormatCellAlignment(row, 1, HorizontalAlignmentValues.Center, VerticalAlignmentValues.Center, true);
                    }
                });
                Assert.Equal(1, saves());
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }
            using var reopened = ExcelDocument.Load(stream);
            for (int row = 1; row <= 20; row++) {
                ExcelCellStyleSnapshot style = reopened.Sheets[0].CellAt(row, 1).GetStyle();
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
                ExcelSheet owner = document.AddWorksheet("Batch owner");
                ExcelSheet target = document.AddWorksheet("Values");
                if (materialized) {
                    target.CellValue(1, 1, "Seed");
                    target.CellBold(1, 1);
                }
                owner.Batch(_ => {
                    target.CellValue(1, 1, "Text");
                    target.CellValue(2, 1, 2D);
                    target.CellValue(3, 1, 3F);
                    target.CellValue(4, 1, 4M);
                    target.CellValue(5, 1, 5);
                    target.CellValue(6, 1, 6L);
                    target.CellValue(7, 1, (short)7);
                    target.CellValue(8, 1, 8U);
                    target.CellValue(9, 1, 9UL);
                    target.CellValue(10, 1, (ushort)10);
                    target.CellValue(11, 1, (byte)11);
                    target.CellValue(12, 1, (sbyte)12);
                    target.CellValue(13, 1, true);
                    target.CellValue(14, 1, new DateTime(2026, 1, 2));
                    target.CellValue(15, 1, TimeSpan.FromHours(36));
                    target.CellValue(16, 1, new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.Zero));
                    target.CellValue<int>(17, 1, 17);
                    target.CellValue(18, 1, (object)18D);
                    target.CellValue(19, 1, 42D);
                    target.CellFormula(19, 1, "6*7");
#if NET6_0_OR_GREATER
                    target.CellValue(20, 1, new DateOnly(2026, 1, 2));
                    target.CellValue(21, 1, new TimeOnly(3, 4, 5));
#endif
                });
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }
            using var reopened = ExcelDocument.Load(stream);
            ExcelSheet saved = reopened.Sheets.Single(sheet => sheet.Name == "Values");
            Assert.Equal("Text", saved.CellAt(1, 1).GetValue<string>());
            for (int row = 2; row <= 12; row++) Assert.Equal((double)row, saved.CellAt(row, 1).GetValue<double>());
            Assert.True(saved.CellAt(13, 1).GetValue<bool>());
            for (int row = 14; row <= 16; row++) Assert.NotEmpty(saved.CellAt(row, 1).GetFormattedText());
            Assert.Equal(17D, saved.CellAt(17, 1).GetValue<double>());
            Assert.Equal(18D, saved.CellAt(18, 1).GetValue<double>());
            Assert.Equal("6*7", saved.GetFormulaText(19, 1));
            Assert.Equal(42D, saved.CellAt(19, 1).GetValue<double>());
#if NET6_0_OR_GREATER
            Assert.NotEmpty(saved.CellAt(20, 1).GetFormattedText());
            Assert.NotEmpty(saved.CellAt(21, 1).GetFormattedText());
#endif
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void BulkMutation_held_wrappers_share_header_and_search_cache_invalidation(bool replaceAll) {
            using var seed = new MemoryStream();
            using (var creator = ExcelDocument.Create(seed)) {
                ExcelSheet sheet = creator.AddWorksheet("Aliases");
                sheet.CellValue(1, 1, "Old");
                sheet.CellValue(1, 2, "Other");
                sheet.CellValue(2, 1, "Value");
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
            else alias.CellValue(1, 1, "New");
            Dictionary<string, int> updated = held.GetHeaderMap();
            Assert.False(updated.ContainsKey("Old"));
            Assert.Equal(1, updated["New"]);
            Assert.Null(held.FindFirst("Old"));
            Assert.Equal("A1", held.FindFirst("New"));
            held.SetByHeader(2, "New", "Written");
            alias.CellValue(1, 1, "Newest");
            Assert.Equal(1, held.GetHeaderMap()["Newest"]);
            Assert.True(originalSnapshot.ContainsKey("Old"));
            Assert.False(originalSnapshot.ContainsKey("Newest"));
            using var destination = new MemoryStream();
            document.Save(destination, new ExcelSaveOptions { ValidateOpenXml = true });
            using var reopened = ExcelDocument.Load(destination);
            Assert.Equal("Newest", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
            Assert.Equal("Written", reopened.Sheets[0].CellAt(2, 1).GetValue<string>());
        }

        [Fact]
        public void BulkMutation_held_wrapper_shared_string_indexes_bind_to_the_restored_table() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                document.SheetCachingEnabled = false;
                ExcelSheet held = document.AddWorksheet("Values");
                ExcelSheet owner = document.AddWorksheet("Transaction");
                held.CellValue(1, 1, "Stable");
                Assert.Throws<InvalidOperationException>(() => owner.ApplyTransactionalMutation(_ => {
                    held.CellValue(2, 1, "Transient");
                    throw new InvalidOperationException("Rollback");
                }, 1, new ExcelMutationPlanOptions(), default));
                owner.CellValue(2, 1, "Replacement");
                held.CellValue(2, 1, "Transient");
                document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = true });
            }
            using var reopened = ExcelDocument.Load(stream);
            Assert.Equal("Transient", reopened.Sheets.Single(sheet => sheet.Name == "Values").CellAt(2, 1).GetValue<string>());
            Assert.Equal("Replacement", reopened.Sheets.Single(sheet => sheet.Name == "Transaction").CellAt(2, 1).GetValue<string>());
        }

        [Fact]
        public void BulkMutation_shared_header_cache_respects_date_options_converters_and_cancellation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Headers");
            sheet.CellValue(1, 1, 45292D);
            sheet.CellValue(2, 1, "Value");
            sheet.FormatCell(1, 1, "yyyy-mm-dd");
            var rawOptions = new ExcelReadOptions { TreatDatesUsingNumberFormat = false };
            Assert.Equal(1, sheet.GetHeaderMap(rawOptions)["45292"]);
            Assert.False(sheet.GetHeaderMap().ContainsKey("45292"));
            Assert.Equal(1, sheet.GetHeaderMap(rawOptions)["45292"]);
            string converted = "Converted";
            var converterOptions = new ExcelReadOptions {
                CellValueConverter = _ => new ExcelCellValue(converted)
            };
            Assert.Equal(1, sheet.GetHeaderMap(converterOptions)["Converted"]);
            converted = "Changed";
            Assert.Equal(1, sheet.GetHeaderMap(converterOptions)["Changed"]);
            Assert.False(sheet.GetHeaderMap().ContainsKey("Changed"));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => {
                sheet.GetHeaderMap(new ExcelReadOptions { CancellationToken = cancellation.Token });
            });
        }

        [Theory]
        [InlineData("cell")]
        [InlineData("range")]
        [InlineData("apply")]
        [InlineData("redefine")]
        public void BulkMutation_header_cache_follows_number_format_and_named_style_changes(string route) {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Headers");
            ExcelSheet source = document.AddWorksheet("Styles");
            sheet.CellValue(1, 1, 45292D);
            sheet.CellValue(2, 1, "Value");
            source.CellValue(1, 1, "Date style");
            source.FormatCell(1, 1, "yyyy-mm-dd");
            source.CellValue(1, 2, "General style");
            source.DefineNamedStyle("Header", 1, route == "redefine" ? 2 : 1);
            if (route == "redefine") sheet.ApplyNamedStyle("Header", "A1:A1");
            Assert.Equal(1, sheet.GetHeaderMap()["45292"]);
            switch (route) {
                case "cell": sheet.FormatCell(1, 1, "yyyy-mm-dd"); break;
                case "range": sheet.FormatRange("A1:A1", "yyyy-mm-dd"); break;
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

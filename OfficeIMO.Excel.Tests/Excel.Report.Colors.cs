using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelReportColorTests {
        [Fact]
        public void ReportColors_PreserveNativeThemeSlotsInSavedWorkbook() {
            using var output = new MemoryStream();
            int listRow;
            int kpiRow;
            int legendRow;
            int extraAccentRow;
            using (var document = ExcelDocument.Create()) {
                var report = document.AddSheet("Summary").Report();
                listRow = report.CurrentRow;
                report.BulletedListWithFill(new[] { "Review" }, OfficeColor.Accent1);
                kpiRow = report.CurrentRow;
                report.KpiRow(new (string Label, object? Value)[] { ("Count", 2) }, labelFill: OfficeColor.Accent2);
                legendRow = report.CurrentRow;
                report.SectionLegend(null, new[] { "Status", "Meaning" },
                    new[] { new[] { "ready", "Reviewed" } },
                    new Dictionary<string, OfficeColor>(StringComparer.OrdinalIgnoreCase) { ["Ready"] = OfficeColor.Accent3 },
                    headerFill: OfficeColor.Accent4);
                extraAccentRow = report.CurrentRow;
                report.KpiRow(new (string Label, object? Value)[] { ("Fifth accent", 5) }, labelFill: OfficeColor.Accent5);
                report.BulletedListWithFill(new[] { "Sixth accent" }, OfficeColor.Accent6);
                report.ConditionalColorScale("B20:B21", OfficeColor.Text, OfficeColor.Background);
                report.ConditionalDataBar("C20:C21", OfficeColor.TextMuted);
                document.Save(output);
            }

            output.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            WorkbookPart workbook = spreadsheet.WorkbookPart!;
            WorksheetPart worksheet = Assert.Single(workbook.WorksheetParts);
            AssertTheme(GetReportFill(workbook, worksheet, "A" + listRow).ForegroundColor!, 4U);
            AssertTheme(GetReportFill(workbook, worksheet, "A" + kpiRow).ForegroundColor!, 5U);
            AssertTheme(GetReportFill(workbook, worksheet, "A" + legendRow).ForegroundColor!, 7U);
            AssertTheme(GetReportFill(workbook, worksheet, "A" + (legendRow + 1)).ForegroundColor!, 6U);
            AssertTheme(GetReportFill(workbook, worksheet, "A" + extraAccentRow).ForegroundColor!, 8U);
            AssertTheme(GetReportFill(workbook, worksheet, "A" + (extraAccentRow + 3)).ForegroundColor!, 9U);
            var scale = Assert.Single(worksheet.Worksheet.Descendants<ColorScale>());
            var scaleColors = scale.Elements<Color>().ToArray();
            Assert.Equal(2, scaleColors.Length);
            AssertTheme(scaleColors[0], 1U);
            AssertTheme(scaleColors[1], 0U);
            var bar = Assert.Single(worksheet.Worksheet.Descendants<DataBar>());
            AssertTheme(Assert.Single(bar.Elements<Color>()), 3U);
        }

        [Fact]
        public void ReportColors_KeepDefaultsSpacingPaletteAndLiteralAlpha() {
            using var output = new MemoryStream();
            OfficeTheme theme = OfficeTheme.Dark.WithPalette(OfficeTheme.Dark.Palette.With(
                success: OfficeColor.FromRgba(0x12, 0x34, 0x56, 0x80),
                warning: OfficeColor.Parse("#234567"), danger: OfficeColor.Parse("#345678")));
            using (var document = ExcelDocument.Create()) {
                var report = document.AddSheet("Summary").Report(new ExcelReportOptions {
                    Theme = theme,
                    KeyFill = OfficeColor.Parse("#123456"),
                    DefaultSpacingRows = 2
                });
                report.BulletedListWithFill(null!, OfficeColor.Accent1);
                report.KpiRow(null!);
                Assert.Equal(1, report.CurrentRow);
                report.BulletedListWithFill(new[] { "Review" }, OfficeColor.FromRgba(0x12, 0x34, 0x56, 0x80));
                Assert.Equal(4, report.CurrentRow);
                report.KpiRow(new (string Label, object? Value)[] { ("Count", 2) }, labelFill: null);
                Assert.Equal(8, report.CurrentRow);
                report.SectionLegend(null, new[] { "Status", "Meaning" },
                    new[] { new[] { "Ready", "Reviewed" } }, headerFill: null);
                Assert.Equal(12, report.CurrentRow);
                report.KpiRow(new (string Label, object? Value)[] { ("Warnings", 1) }, labelFill: OfficeColor.Warning);
                report.ConditionalColorScale("B20:B21", OfficeColor.Danger, OfficeColor.Success);
                report.ConditionalDataBar("C20:C21", OfficeColor.FromRgba(0x12, 0x34, 0x56, 0x80));
                document.Save(output);
            }

            output.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            WorkbookPart workbook = spreadsheet.WorkbookPart!;
            WorksheetPart worksheet = Assert.Single(workbook.WorksheetParts);
            Assert.Equal("80123456", GetReportFill(workbook, worksheet, "A1").ForegroundColor!.Rgb!.Value);
            Assert.Equal("FF123456", GetReportFill(workbook, worksheet, "A4").ForegroundColor!.Rgb!.Value);
            Assert.Equal("FF123456", GetReportFill(workbook, worksheet, "A8").ForegroundColor!.Rgb!.Value);
            Assert.Equal(theme.Palette.Warning.ToArgbHex(), GetReportFill(workbook, worksheet, "A12").ForegroundColor!.Rgb!.Value);
            var scale = Assert.Single(worksheet.Worksheet.Descendants<ColorScale>());
            Assert.Equal(new[] { theme.Palette.Danger.ToArgbHex(), theme.Palette.Success.ToArgbHex() },
                scale.Elements<Color>().Select(color => color.Rgb!.Value).ToArray());
            var bar = Assert.Single(worksheet.Worksheet.Descendants<DataBar>());
            Assert.Equal("80123456", Assert.Single(bar.Elements<Color>()).Rgb!.Value);
        }

        [Fact]
        public void ReportColors_KeepThemeKeyFillAndDictionaryComparer() {
            using var output = new MemoryStream();
            int legendRow;
            using (var document = ExcelDocument.Create()) {
                var report = document.AddSheet("Summary").Report(new ExcelReportOptions {
                    Theme = OfficeTheme.Dark,
                    KeyFill = OfficeColor.Accent5
                });
                report.KpiRow(new (string Label, object? Value)[] { ("Omitted fill", 1) });
                report.KpiRow(new (string Label, object? Value)[] { ("Null fill", 2) }, labelFill: null);
                legendRow = report.CurrentRow;
                report.SectionLegend(null, new[] { "Status" }, new[] { new[] { "ready" }, new[] { "Ready" } },
                    new Dictionary<string, OfficeColor>(StringComparer.Ordinal) { ["Ready"] = OfficeColor.Success },
                    headerFill: null);
                document.Save(output);
            }

            output.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            WorkbookPart workbook = spreadsheet.WorkbookPart!;
            WorksheetPart worksheet = Assert.Single(workbook.WorksheetParts);
            AssertTheme(GetReportFill(workbook, worksheet, "A1").ForegroundColor!, 8U);
            AssertTheme(GetReportFill(workbook, worksheet, "A4").ForegroundColor!, 8U);
            AssertTheme(GetReportFill(workbook, worksheet, "A" + legendRow).ForegroundColor!, 8U);
            Assert.Null(GetReportFill(workbook, worksheet, "A" + (legendRow + 1)).ForegroundColor);
            Assert.Equal(OfficeTheme.Dark.Palette.Success.ToArgbHex(),
                GetReportFill(workbook, worksheet, "A" + (legendRow + 2)).ForegroundColor!.Rgb!.Value);
        }

        [Fact]
        public void SheetColors_RequireResolvedStatusRolesWithCustomWorkbookTheme() {
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Create()) {
                var sheet = document.AddSheet("Summary");
                document.EnsureWorkbookTheme();
                var themeXml = XDocument.Parse(document.GetWorkbookThemeXml()!);
                XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
                themeXml.Descendants(drawing + "accent1").Single().ReplaceNodes(
                    new XElement(drawing + "srgbClr", new XAttribute("val", "102030")));
                document.SetWorkbookThemeXml(themeXml.ToString());

                // DrawingML has no status slots. Only the report has a palette for these roles.
                foreach (OfficeColor role in new[] { OfficeColor.Success, OfficeColor.Warning, OfficeColor.Danger }) {
                    Assert.Throws<InvalidOperationException>(() => sheet.CellBackground(99, 1, role));
                    Assert.Throws<InvalidOperationException>(() => sheet.AddConditionalColorScale("B99:B100", OfficeColor.Accent1, role));
                    Assert.Throws<InvalidOperationException>(() => sheet.AddConditionalDataBar("C99:C100", role));
                }
                sheet.CellBackground(1, 1, OfficeColor.Accent1);
                Assert.Equal("FF102030", document.ResolveThemeColorArgb(4U));
                document.Save(output);
            }

            output.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(output, false);
            WorkbookPart workbook = spreadsheet.WorkbookPart!;
            WorksheetPart worksheet = Assert.Single(workbook.WorksheetParts);
            AssertTheme(GetReportFill(workbook, worksheet, "A1").ForegroundColor!, 4U);
            Assert.Single(worksheet.Worksheet.Descendants<Cell>());
            Assert.Empty(worksheet.Worksheet.Descendants<ConditionalFormatting>());
        }

        private static PatternFill GetReportFill(WorkbookPart workbook, WorksheetPart worksheet, string reference) {
            var cell = worksheet.Worksheet.Descendants<Cell>().Single(value => value.CellReference?.Value == reference);
            Stylesheet styles = workbook.WorkbookStylesPart!.Stylesheet;
            CellFormat format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)(cell.StyleIndex?.Value ?? 0U));
            return styles.Fills!.Elements<Fill>().ElementAt((int)format.FillId!.Value).PatternFill!;
        }

        private static void AssertTheme(ColorType color, uint index) {
            Assert.Equal(index, color.Theme?.Value);
            Assert.Null(color.Rgb);
        }
    }
}

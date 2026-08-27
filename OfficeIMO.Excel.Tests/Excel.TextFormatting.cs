using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelTextFormattingTests {
    [Fact]
    public void CellAndRangeApisPersistNativeFontStyles() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Text");
                sheet.CellAt(1, 1).SetValue("Primary")
                    .SetBold().SetItalic()
                    .SetUnderline(ExcelUnderlineStyle.DoubleAccounting)
                    .SetStrikethrough().SetSuperscript()
                    .SetFontName("Aptos").SetFontSize(14).SetFontColor("336699");
                sheet.Range("A2:A3").SetUnderline(ExcelUnderlineStyle.Double)
                    .SetItalic().SetStrikethrough().SetSubscript();
                document.Save();
            }

            using SpreadsheetDocument package = SpreadsheetDocument.Open(path, false);
            WorkbookPart workbook = package.WorkbookPart!;
            WorksheetPart worksheet = workbook.WorksheetParts.Single();
            Stylesheet stylesheet = workbook.WorkbookStylesPart!.Stylesheet!;

            Font primary = ResolveFont(worksheet, stylesheet, "A1");
            Assert.NotNull(primary.Bold);
            Assert.NotNull(primary.Italic);
            Assert.Equal(UnderlineValues.DoubleAccounting, primary.Underline!.Val!.Value);
            Assert.NotNull(primary.Strike);
            Assert.Equal(VerticalAlignmentRunValues.Superscript, primary.VerticalTextAlignment!.Val!.Value);
            Assert.Equal("Aptos", primary.FontName!.Val!.Value);
            Assert.Equal(14D, primary.FontSize!.Val!.Value);
            Assert.Equal("FF336699", primary.Color!.Rgb!.Value);

            foreach (string reference in new[] { "A2", "A3" }) {
                Font ranged = ResolveFont(worksheet, stylesheet, reference);
                Assert.NotNull(ranged.Italic);
                Assert.Equal(UnderlineValues.Double, ranged.Underline!.Val!.Value);
                Assert.NotNull(ranged.Strike);
                Assert.Equal(VerticalAlignmentRunValues.Subscript, ranged.VerticalTextAlignment!.Val!.Value);
            }
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void CellAndRangeCaseTransformsPreserveFormattingAndSkipNonTextValues() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Text");
                sheet.CellAt(1, 1).SetValue("mixed CASE").SetBold().TransformTextCase(OfficeTextCase.TitleCase);
                sheet.CellAt(2, 1).SetRichText(
                    new ExcelRichTextRun("Styled") { Bold = true, UnderlineStyle = ExcelUnderlineStyle.Double },
                    new ExcelRichTextRun(" TEXT") { Italic = true, VerticalTextAlignment = ExcelVerticalTextAlignment.Subscript });
                sheet.CellAt(3, 1).SetValue(42);
                sheet.CellAt(4, 1).SetFormula("A3+1");

                sheet.Range("A2:A4").TransformTextCase(OfficeTextCase.ToggleCase);
                document.Save();
            }

            using ExcelDocument reloaded = ExcelDocument.Load(path);
            ExcelSheet reloadedSheet = reloaded.Sheets[0];
            Assert.Equal("Mixed Case", reloadedSheet.CellAt(1, 1).GetValue<string>());
            ExcelRichTextRun[] runs = reloadedSheet.CellAt(2, 1).GetRichText().ToArray();
            Assert.Equal(new[] { "sTYLED", " text" }, runs.Select(run => run.Text));
            Assert.True(runs[0].Bold);
            Assert.Equal(ExcelUnderlineStyle.Double, runs[0].UnderlineStyle);
            Assert.True(runs[1].Italic);
            Assert.Equal(ExcelVerticalTextAlignment.Subscript, runs[1].VerticalTextAlignment);
            Assert.Equal(42D, reloadedSheet.CellAt(3, 1).GetValue<double>());
            Assert.Equal("A3+1", reloadedSheet.CellAt(4, 1).GetValue().Formula);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static Font ResolveFont(WorksheetPart worksheet, Stylesheet stylesheet, string reference) {
        Cell cell = worksheet.Worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == reference);
        CellFormat format = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt(checked((int)cell.StyleIndex!.Value));
        return stylesheet.Fonts!.Elements<Font>().ElementAt(checked((int)format.FontId!.Value));
    }
}

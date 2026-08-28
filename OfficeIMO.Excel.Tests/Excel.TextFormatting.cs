using System;
using System.Globalization;
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

    [Fact]
    public void FormulaWithCachedStringResultIsNeverRewrittenByCaseTransforms() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                document.AddWorksheet("Text").CellAt(1, 1).SetFormula("\"mixed Case\"");
                document.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                Cell cell = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single();
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue("mixed Case");
                package.WorkbookPart.WorksheetParts.Single().Worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(path)) {
                ExcelSheet sheet = document.Sheets[0];
                Assert.True(sheet.TryGetCellValueSnapshot(1, 1, out ExcelCellValueSnapshot? snapshot));
                Assert.Equal(ExcelCellValueKind.Formula, snapshot!.Kind);
                Assert.False(sheet.TransformCellTextCase(1, 1, OfficeTextCase.Uppercase));
                document.Save();
            }

            using SpreadsheetDocument verified = SpreadsheetDocument.Open(path, false);
            Cell actual = verified.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single();
            Assert.Equal("\"mixed Case\"", actual.CellFormula!.Text);
            Assert.Equal("mixed Case", actual.CellValue!.Text);
            Assert.Equal(CellValues.String, actual.DataType!.Value);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void RichTextCaseTransformsDoNotResetContextAtRunBoundaries() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Text");
        sheet.CellAt(1, 1).SetRichText(
            new ExcelRichTextRun("hELLO. ") { Bold = true },
            new ExcelRichTextRun("aNOTHER") { Italic = true },
            new ExcelRichTextRun(" SENTENCE") { Underline = true });

        sheet.CellAt(1, 1).TransformTextCase(OfficeTextCase.SentenceCase, CultureInfo.InvariantCulture);

        ExcelRichTextRun[] actual = sheet.CellAt(1, 1).GetRichText().ToArray();
        Assert.Equal(new[] { "Hello. ", "Another", " sentence" }, actual.Select(run => run.Text));
        Assert.True(actual[0].Bold);
        Assert.True(actual[1].Italic);
        Assert.True(actual[2].Underline);
    }

    [Fact]
    public void RichTextCaseTransformsPreserveNativeRunProperties() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        string expectedProperties;
        try {
            using (ExcelDocument created = ExcelDocument.Create(path)) {
                created.AddWorksheet("Text").CellAt(1, 1).SetRichText(
                    new ExcelRichTextRun("istanbul") { Bold = true, FontName = "Aptos" });
                created.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                RunProperties properties = package.WorkbookPart!.WorksheetParts.Single().Worksheet
                    .Descendants<Cell>().Single(item => item.CellReference?.Value == "A1")
                    .InlineString!.Elements<Run>().Single().RunProperties!;
                properties.AddChild(new Color { Theme = 4U, Tint = 0.25D }, true);
                properties.AddChild(new FontScheme { Val = FontSchemeValues.Minor }, true);
                expectedProperties = properties.OuterXml;
            }

            using (ExcelDocument loaded = ExcelDocument.Load(path)) {
                loaded.Sheets[0].CellAt(1, 1)
                    .TransformTextCase(OfficeTextCase.TitleCase, CultureInfo.GetCultureInfo("tr-TR"));
                loaded.Save();
            }

            using SpreadsheetDocument verified = SpreadsheetDocument.Open(path, false);
            Run run = verified.WorkbookPart!.WorksheetParts.Single().Worksheet
                .Descendants<Cell>().Single(item => item.CellReference?.Value == "A1")
                .InlineString!.Elements<Run>().Single();
            Assert.Equal("İstanbul", run.Text!.Text);
            Assert.Equal(expectedProperties, run.RunProperties!.OuterXml);
            Assert.Equal(4U, run.RunProperties.GetFirstChild<Color>()!.Theme!.Value);
            Assert.Equal(FontSchemeValues.Minor, run.RunProperties.GetFirstChild<FontScheme>()!.Val!.Value);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SharedRichTextCaseTransformsCloneNativePropertiesWithoutChangingSiblingCells() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        string expectedProperties;
        try {
            using (ExcelDocument created = ExcelDocument.Create(path)) {
                ExcelSheet sheet = created.AddWorksheet("Text");
                sheet.CellAt(1, 1).SetValue("First");
                sheet.CellAt(2, 1).SetValue("Second");
                created.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorkbookPart workbook = package.WorkbookPart!;
                SharedStringTablePart sharedPart = workbook.SharedStringTablePart ?? workbook.AddNewPart<SharedStringTablePart>();
                SharedStringTable shared = sharedPart.SharedStringTable ??= new SharedStringTable();
                var properties = new RunProperties();
                properties.AddChild(new RunFont { Val = "+mn-lt" }, true);
                properties.AddChild(new Color { Theme = 5U, Tint = -0.2D }, true);
                properties.AddChild(new FontScheme { Val = FontSchemeValues.Minor }, true);
                var item = new SharedStringItem(new Run(properties, new Text("istanbul")));
                shared.Append(item);
                int sharedIndex = shared.Elements<SharedStringItem>().Count() - 1;
                expectedProperties = properties.OuterXml;

                Worksheet worksheet = workbook.WorksheetParts.Single().Worksheet;
                foreach (string reference in new[] { "A1", "A2" }) {
                    Cell cell = worksheet.Descendants<Cell>().Single(candidate => candidate.CellReference?.Value == reference);
                    cell.InlineString = null;
                    cell.DataType = CellValues.SharedString;
                    cell.CellValue = new CellValue(sharedIndex.ToString(CultureInfo.InvariantCulture));
                }
                worksheet.Save();
                shared.Save();
            }

            using (ExcelDocument loaded = ExcelDocument.Load(path)) {
                loaded.Sheets[0].CellAt(1, 1)
                    .TransformTextCase(OfficeTextCase.TitleCase, CultureInfo.GetCultureInfo("tr-TR"));
                loaded.Save();
            }

            using SpreadsheetDocument verified = SpreadsheetDocument.Open(path, false);
            WorkbookPart verifiedWorkbook = verified.WorkbookPart!;
            Worksheet verifiedWorksheet = verifiedWorkbook.WorksheetParts.Single().Worksheet;
            Cell transformedCell = verifiedWorksheet.Descendants<Cell>()
                .Single(candidate => candidate.CellReference?.Value == "A1");
            Cell siblingCell = verifiedWorksheet.Descendants<Cell>()
                .Single(candidate => candidate.CellReference?.Value == "A2");
            Run transformedRun = transformedCell.InlineString!.Elements<Run>().Single();
            Assert.Equal(CellValues.InlineString, transformedCell.DataType!.Value);
            Assert.Equal("İstanbul", transformedRun.Text!.Text);
            Assert.Equal(expectedProperties, transformedRun.RunProperties!.OuterXml);
            Assert.Equal(CellValues.SharedString, siblingCell.DataType!.Value);
            int siblingIndex = int.Parse(siblingCell.CellValue!.InnerText, CultureInfo.InvariantCulture);
            SharedStringItem siblingItem = verifiedWorkbook.SharedStringTablePart!.SharedStringTable!
                .Elements<SharedStringItem>().ElementAt(siblingIndex);
            Assert.Equal("istanbul", siblingItem.Elements<Run>().Single().Text!.Text);
            Assert.Equal(expectedProperties, siblingItem.Elements<Run>().Single().RunProperties!.OuterXml);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ExplicitlyDisabledOpenXmlFontPropertiesRemainDisabledAcrossStyleSnapshots() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument created = ExcelDocument.Create(path)) {
                created.AddWorksheet("Text").CellAt(1, 1).SetValue("Plain")
                    .SetBold().SetItalic().SetUnderline().SetStrikethrough();
                created.Sheets[0].CellAt(2, 1).SetRichText(
                    new ExcelRichTextRun("Rich") {
                        Bold = true,
                        Italic = true,
                        Underline = true,
                        Strikethrough = true,
                    });
                created.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorkbookPart workbook = package.WorkbookPart!;
                WorksheetPart worksheet = workbook.WorksheetParts.Single();
                Stylesheet stylesheet = workbook.WorkbookStylesPart!.Stylesheet!;
                Font font = ResolveFont(worksheet, stylesheet, "A1");
                font.Bold = new Bold { Val = false };
                font.Italic = new Italic { Val = false };
                font.Underline = new Underline { Val = UnderlineValues.None };
                font.Strike = new Strike { Val = false };

                Cell richCell = worksheet.Worksheet.Descendants<Cell>()
                    .Single(item => item.CellReference?.Value == "A2");
                RunProperties runProperties = richCell.InlineString!.Elements<Run>().Single().RunProperties!;
                runProperties.GetFirstChild<Bold>()!.Val = false;
                runProperties.GetFirstChild<Italic>()!.Val = false;
                runProperties.GetFirstChild<Underline>()!.Val = UnderlineValues.None;
                runProperties.GetFirstChild<Strike>()!.Val = false;
                stylesheet.Save();
            }

            using ExcelDocument loaded = ExcelDocument.Load(path);
            ExcelCellStyleSnapshot direct = loaded.Sheets[0].CellAt(1, 1).GetStyle();
            ExcelWorksheetSnapshot worksheetSnapshot = Assert.Single(loaded.CreateInspectionSnapshot().Worksheets);
            ExcelCellStyleSnapshot inspected = worksheetSnapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style!;
            foreach (ExcelCellStyleSnapshot style in new[] { direct, inspected }) {
                Assert.False(style.Bold);
                Assert.False(style.Italic);
                Assert.False(style.Underline);
                Assert.Equal(ExcelUnderlineStyle.None, style.UnderlineStyle);
                Assert.False(style.Strikethrough);
            }

            ExcelRichTextRun directRun = Assert.Single(loaded.Sheets[0].CellAt(2, 1).GetRichText());
            ExcelRichTextRun inspectedRun = Assert.Single(
                worksheetSnapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1).RichTextRuns);
            foreach (ExcelRichTextRun run in new[] { directRun, inspectedRun }) {
                Assert.False(run.Bold);
                Assert.False(run.Italic);
                Assert.False(run.Underline);
                Assert.Equal(ExcelUnderlineStyle.None, run.UnderlineStyle);
                Assert.False(run.Strikethrough);
            }
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

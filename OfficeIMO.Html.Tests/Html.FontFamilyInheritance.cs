using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlFontFamilyInheritanceTests {
    private const string Marker = "FAMILYMARKER";

    [Theory]
    [InlineData("excel", "font-family:inherit", false)]
    [InlineData("excel", "font-family:unset", false)]
    [InlineData("excel", "font:inherit", false)]
    [InlineData("excel", "font-family:var(--missing-family, inherit)", false)]
    [InlineData("excel", "font-family:'inherit'", true)]
    [InlineData("onenote", "font-family:inherit", false)]
    [InlineData("onenote", "font-family:unset", false)]
    [InlineData("onenote", "font:inherit", false)]
    [InlineData("onenote", "font-family:var(--missing-family, inherit)", false)]
    [InlineData("onenote", "font-family:'inherit'", true)]
    public void HtmlFontFamilyInheritance_ReachesEditableTargetAndSavedArtifact(
        string target, string declaration, bool namedFace) {
        const string heading = "<h1>Font inheritance controls</h1>";
        string baselineHtml = heading + "<p><span>" + Marker + "</span></p>";
        string actualHtml = heading + "<p><span style=\"" + declaration + "\">" + Marker + "</span></p>";

        (string?[] Imported, string?[] Reopened) baseline = ReadTargetFonts(target, baselineHtml);
        (string?[] Imported, string?[] Reopened) actual = ReadTargetFonts(target, actualHtml);

        // Every format keeps its own native default; no common font is assumed.
        // A quoted family called "inherit" is an authored face, not a CSS keyword.
        Assert.NotEmpty(actual.Imported);
        Assert.NotEmpty(actual.Reopened);
        if (namedFace) {
            Assert.Contains("inherit", actual.Imported);
            Assert.Contains("inherit", actual.Reopened);
        } else {
            Assert.Equal(baseline.Imported, actual.Imported);
            Assert.Equal(baseline.Reopened, actual.Reopened);
        }
    }

    private static (string?[] Imported, string?[] Reopened) ReadTargetFonts(string target, string html) {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        switch (target) {
            case "excel": {
                using ExcelDocument workbook = source.ToExcelDocumentResult(
                    new HtmlToExcelOptions { Mode = HtmlImportMode.Generic }).RequireValue();
                string? imported = ExcelFont(workbook);
                using var artifact = new MemoryStream();
                workbook.Save(artifact);
                using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(artifact.ToArray()));
                return (new[] { imported }, new[] { ExcelFont(reopened) });
            }
            case "onenote": {
                OneNoteSection section = source.ToOneNoteSectionResult().RequireValue();
                string? imported = OneNoteFont(section);
                using var artifact = new MemoryStream(OneNoteSectionWriter.Write(section));
                OneNoteSection reopened = OneNoteSectionReader.Read(artifact);
                return (new[] { imported }, new[] { OneNoteFont(reopened) });
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(target));
        }
    }

    private static string? ExcelFont(ExcelDocument workbook) {
        ExcelSheet sheet = Assert.Single(workbook.Sheets);
        int row = Assert.Single(Enumerable.Range(1, 4), index =>
            sheet.CellAt(index, 1).GetValue<string>() == Marker);
        var runs = sheet.CellAt(row, 1).GetRichText();
        string? cellFont = sheet.GetCellStyle(row, 1).FontName;
        return runs.Count == 0 ? cellFont
            : Assert.Single(runs, run => run.Text == Marker).FontName ?? cellFont;
    }

    private static string? OneNoteFont(OneNoteSection section) {
        OneNoteTextRun run = Assert.Single(section.Pages
            .SelectMany(page => page.Outlines)
            .SelectMany(outline => outline.Children)
            .OfType<OneNoteParagraph>()
            .SelectMany(paragraph => paragraph.Runs), candidate => candidate.Text == Marker);
        return run.Style.FontFamily;
    }
}

using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("font-weight:300", "<strong>X</strong>", false, false)]
    [InlineData("font-weight:300", "<b><span>X</span></b>", false, false)]
    [InlineData("font-weight:300", "<strong style='font-weight:inherit'>X</strong>", false, false)]
    [InlineData("font-weight:400", "<strong style='font-weight:unset'>X</strong>", false, false)]
    [InlineData("font-weight:900", "<strong><span style='font-weight:initial'>X</span></strong>", false, false)]
    [InlineData("font-weight:400", "<strong>X</strong>", true, false)]
    [InlineData("font-style:normal", "<em style='font-style:inherit'>X</em>", false, false)]
    [InlineData("font-style:normal", "<i style='font-style:unset'>X</i>", false, false)]
    [InlineData("font-style:italic", "<em><span style='font-style:initial'>X</span></em>", false, false)]
    [InlineData("font-style:normal", "<cite><span>X</span></cite>", false, true)]
    public void SemanticFontDefaultsReachEditableRichText(
        string parentStyle, string markup, bool expectedBold, bool expectedItalic) {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<body style='font:16px Arial;" + parentStyle + "'><p>" + markup + "</p></body>");
        HtmlSemanticRun run = Assert.Single(Assert.Single(Assert.Single(source.SemanticDocument.Sections).Blocks).Runs);
        Assert.Equal("X", run.Text);
        Assert.Equal(expectedBold, run.Bold);
        Assert.Equal(expectedItalic, run.Italic);

        using ExcelDocument workbook = source.ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic }).RequireValue();
        ExcelSheet sheet = Assert.Single(workbook.Sheets);
        ExcelCell cell = Assert.Single(Enumerable.Range(1, 4).Select(row => sheet.CellAt(row, 1)),
            candidate => candidate.GetValue<string>() == "X");
        ExcelRichTextRun exported = Assert.Single(cell.GetRichText());
        Assert.Equal(expectedBold, exported.Bold);
        Assert.Equal(expectedItalic, exported.Italic);
    }

    [Theory]
    [InlineData("strong", "text", "font-weight:initial", 400, OfficeFontSlant.Normal)]
    [InlineData("strong", "g", "font-weight:initial", 400, OfficeFontSlant.Normal)]
    [InlineData("strong", "tspan", "font-weight:initial", 400, OfficeFontSlant.Normal)]
    [InlineData("em", "text", "font-style:initial", 400, OfficeFontSlant.Normal)]
    [InlineData("em", "g", "font:16px Arial", 400, OfficeFontSlant.Normal)]
    [InlineData("em", "tspan", "all:initial", 400, OfficeFontSlant.Normal)]
    [InlineData("strong", "text", "font-weight:inherit", 700, OfficeFontSlant.Normal)]
    [InlineData("em", "text", "font-style:unset", 400, OfficeFontSlant.Italic)]
    [InlineData("strong", "g", "font-weight:revert", 700, OfficeFontSlant.Normal)]
    [InlineData("em", "tspan", "font-style:revert", 400, OfficeFontSlant.Italic)]
    public void SemanticFontDefaultsPreserveInlineSvgResetsAndInheritance(
        string parent, string carrier, string style, int expectedWeight, OfficeFontSlant expectedSlant) {
        string text = carrier switch {
            "g" => "<g style='" + style + "'><text x='0' y='20'>X</text></g>",
            "tspan" => "<text x='0' y='20'><tspan style='" + style + "'>X</tspan></text>",
            _ => "<text x='0' y='20' style='" + style + "'>X</text>"
        };
        OfficeDrawingText run = Assert.Single(InlineSvgTextRuns(
            "<body style='font:16px Arial'><" + parent + "><svg width='100' height='30'>"
            + text + "</svg></" + parent + "></body>"));

        Assert.Equal("X", run.Text);
        Assert.Equal(expectedWeight, run.Font.Face.Weight);
        Assert.Equal(expectedSlant, run.Font.Face.Slant);
    }
}

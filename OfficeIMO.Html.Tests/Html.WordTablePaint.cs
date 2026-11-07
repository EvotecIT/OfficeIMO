using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class WordHtmlTablePaintTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeHeaderSemanticsAndZeroSpacingDoNotDependOnDefaultCss(bool includeCss) {
        using var document = WordDocument.Create();
        var table = document.AddTable(2, 2, WordTableStyle.TableNormal);
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Item";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Count";
        table.Rows[1].Cells[0].Paragraphs[0].Text = "Source";
        table.Rows[1].Cells[1].Paragraphs[0].Text = "12";
        string before = SourceXml(document);

        var html = new HtmlParser().ParseDocument(document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = includeCss }));
        Assert.Equal(new[] { "Item", "Count" }, html.QuerySelectorAll("thead th[scope=col]").Select(cell => cell.TextContent));
        Assert.Empty(html.QuerySelectorAll("tbody td[scope]"));
        string style = html.QuerySelector("table")!.GetAttribute("style")!;
        Assert.Contains("border-collapse:collapse", style);
        Assert.Contains("border-spacing:0", style);
        Assert.Equal(before, SourceXml(document));

        table._tableProperties!.TableCellSpacing = new TableCellSpacing { Type = TableWidthUnitValues.Dxa, Width = "100" };
        before = SourceXml(document);
        var spaced = new HtmlParser().ParseDocument(document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = includeCss }));
        string spacedStyle = spaced.QuerySelector("table")!.GetAttribute("style")!;
        Assert.Contains("border-collapse:separate", spacedStyle);
        Assert.Contains("border-spacing:5pt", spacedStyle);
        Assert.Equal(2, spaced.QuerySelectorAll("thead th[scope=col]").Length);
        Assert.Equal(before, SourceXml(document));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InheritedTableBackdropStaysSeparateFromCellNoFill(bool includeCss) {
        var theme = OfficeTheme.Modern.WithPalette(OfficeTheme.Modern.Palette.With(accent1: OfficeColor.Parse("#123456")));
        using var document = WordDocument.Create(new WordCreateOptions { Theme = theme });
        var table = document.AddTable(1, 1, WordTableStyle.TableNormal);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Backdrop";
        table._tableProperties!.TableCellSpacing = new TableCellSpacing { Type = TableWidthUnitValues.Dxa, Width = "100" };
        document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.Append(
            new Style(new StyleTableProperties(new Shading {
                Val = ShadingPatternValues.Clear, Fill = "FF0000", ThemeFill = ThemeColorValues.Accent1
            })) { Type = StyleValues.Table, StyleId = "BackdropBase" },
            new Style(new BasedOn { Val = "BackdropBase" }, new StyleTableCellProperties(new Shading {
                Val = ShadingPatternValues.Clear, Fill = "CC3300"
            })) { Type = StyleValues.Table, StyleId = "BackdropChild" });
        table.SetStyleId("BackdropChild");
        table.Rows[0].Cells[0]._tableCell.TableCellProperties!.Shading = new Shading {
            Val = ShadingPatternValues.Nil, Fill = "0000FF"
        };
        Assert.Empty(document.ValidateDocument());
        string before = SourceXml(document);

        var inherited = new HtmlParser().ParseDocument(document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = includeCss }));
        Assert.Contains("background-color:#123456", inherited.QuerySelector("table")!.GetAttribute("style"));
        Assert.Contains("background-color:transparent", inherited.QuerySelector("td")!.GetAttribute("style"));
        Assert.Equal(before, SourceXml(document));

        table._tableProperties.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = "445566" };
        before = SourceXml(document);
        var direct = new HtmlParser().ParseDocument(document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = includeCss }));
        Assert.Contains("background-color:#445566", direct.QuerySelector("table")!.GetAttribute("style"));
        Assert.Contains("background-color:transparent", direct.QuerySelector("td")!.GetAttribute("style"));
        Assert.Equal(before, SourceXml(document));

        table._tableProperties.Shading = new Shading { Val = ShadingPatternValues.Nil, Fill = "FF0000" };
        before = SourceXml(document);
        var cleared = new HtmlParser().ParseDocument(document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = includeCss }));
        Assert.Contains("background-color:transparent", cleared.QuerySelector("table")!.GetAttribute("style"));
        Assert.Contains("background-color:transparent", cleared.QuerySelector("td")!.GetAttribute("style"));
        Assert.Equal(before, SourceXml(document));
    }

    [Fact]
    public void UnsupportedSourceBorderPatternsAreReportedAndCanRejectLoss() {
        using var document = WordDocument.Create();
        var table = document.AddTable(1, 1, WordTableStyle.TableNormal);
        document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.Append(new Style(
            new StyleTableCellProperties(new TableCellBorders(new TopBorder {
                Val = BorderValues.DotDash, Size = 8U, Color = "123456"
            }))) { Type = StyleValues.Table, StyleId = "BorderPattern" });
        table.SetStyleId("BorderPattern");

        string html = document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = false }, out var report);
        var diagnostic = Assert.Single(report.Diagnostics, item => item.Code == "word-table-border-pattern-approximated");
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Contains("border-top:1pt solid #123456", html);
        Assert.Throws<OfficeConversionException>(() => document.ToHtml(new WordHtmlExportOptions { FailOnLoss = true }));

        table.Rows[0].Cells[0].Borders.TopStyle = WordBorderStyle.Single;
        document.ToHtml(null, out var clearedReport);
        Assert.DoesNotContain(clearedReport.Diagnostics, item => item.Code == "word-table-border-pattern-approximated");
    }

    [Fact]
    public void SourceDiagonalBordersAreReportedUnlessExplicitlyCleared() {
        using var document = WordDocument.Create();
        var table = document.AddTable(1, 1, WordTableStyle.TableNormal);
        document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.Append(new Style(
            new StyleTableCellProperties(new TableCellBorders(new TopLeftToBottomRightCellBorder {
                Val = BorderValues.Single, Size = 8U, Color = "123456"
            }))) { Type = StyleValues.Table, StyleId = "DiagonalBorder" });
        table.SetStyleId("DiagonalBorder");
        var cell = table.Rows[0].Cells[0];
        cell.Borders.TopRightToBottomLeftStyle = WordBorderStyle.Single;

        document.ToHtml(null, out var report);
        var diagnostic = Assert.Single(report.Diagnostics, item => item.Code == "word-table-diagonal-border-omitted");
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.Throws<OfficeConversionException>(() => document.ToHtml(new WordHtmlExportOptions { FailOnLoss = true }));

        cell.Borders.TopLeftToBottomRightStyle = WordBorderStyle.Nil;
        cell.Borders.TopRightToBottomLeftStyle = WordBorderStyle.None;
        document.ToHtml(null, out var clearedReport);
        Assert.DoesNotContain(clearedReport.Diagnostics, item => item.Code == "word-table-diagonal-border-omitted");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TablePaintUsesSourceInheritanceConditionsAndThemeEvenWithoutDefaultCss(bool includeCss) {
        var theme = OfficeTheme.Modern.WithPalette(OfficeTheme.Modern.Palette.With(accent1: OfficeColor.Parse("#123456")));
        using var document = WordDocument.Create(new WordCreateOptions { Theme = theme });
        var table = document.AddTable(5, 2, WordTableStyle.TableNormal);
        table.ConditionalFormattingFirstRow = true;
        table.ConditionalFormattingLastRow = true;
        table.ConditionalFormattingFirstColumn = true;
        table.ConditionalFormattingNoHorizontalBand = false;
        table.ConditionalFormattingNoVerticalBand = true;
        table.RepeatHeaderRowAtTheTopOfEachPage = false;
        foreach (var cell in table.Rows.SelectMany(row => row.Cells)) cell.Paragraphs[0].Text = "Source";

        var baseStyle = new Style(new StyleTableCellProperties(
            new Shading { Val = ShadingPatternValues.Clear, Fill = "ABABAB" })) { Type = StyleValues.Table, StyleId = "BasePaint" };
        var derived = new Style(new BasedOn { Val = "BasePaint" },
            new TableStyleProperties(new TableStyleConditionalFormattingTableCellProperties(
                new TableCellBorders(new BottomBorder { Val = BorderValues.Single, Size = 2U, Color = "00FF00", ThemeColor = ThemeColorValues.Accent1 }),
                new Shading { Val = ShadingPatternValues.Clear, Fill = "FF0000", ThemeFill = ThemeColorValues.Accent1 })) { Type = TableStyleOverrideValues.FirstRow },
            new TableStyleProperties(new TableStyleConditionalFormattingTableCellProperties(
                new Shading { Val = ShadingPatternValues.Clear, Fill = "0000FF" })) { Type = TableStyleOverrideValues.LastRow },
            new TableStyleProperties(new TableStyleConditionalFormattingTableCellProperties(
                new Shading { Val = ShadingPatternValues.Clear, Fill = "00FF00" })) { Type = TableStyleOverrideValues.Band1Horizontal },
            new TableStyleProperties(new TableStyleConditionalFormattingTableCellProperties(
                new Shading { Val = ShadingPatternValues.Clear, Fill = "222222" })) { Type = TableStyleOverrideValues.FirstColumn }) {
            Type = StyleValues.Table, StyleId = "DerivedPaint"
        };
        document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.Append(baseStyle, derived);
        table.SetStyleId("DerivedPaint");
        table.Rows[1].Cells[0]._tableCell.TableCellProperties!.Append(new Shading { Val = ShadingPatternValues.Nil, Fill = "FF0000" });
        table.Rows[3].Cells[0]._tableCell.TableCellProperties!.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = "auto" });

        var html = new HtmlParser().ParseDocument(document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = includeCss }));
        var rows = html.QuerySelectorAll("tr").ToArray();
        string CellStyle(int row, int column) => rows[row].Children[column].GetAttribute("style")!;

        Assert.Contains("background-color:#123456", CellStyle(0, 0));
        Assert.Contains("border-bottom:0.25pt solid #123456", CellStyle(0, 0));
        Assert.Contains("background-color:transparent", CellStyle(1, 0));
        Assert.Contains("background-color:#00FF00", CellStyle(1, 1));
        Assert.Contains("background-color:#222222", CellStyle(2, 0));
        Assert.Contains("background-color:#ABABAB", CellStyle(2, 1));
        Assert.Contains("background-color:transparent", CellStyle(3, 0));
        Assert.Contains("background-color:#00FF00", CellStyle(3, 1));
        Assert.Contains("background-color:#0000FF", CellStyle(4, 0));

        table._tableProperties!.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = "333333" };
        var directTableHtml = new HtmlParser().ParseDocument(document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = includeCss }));
        Assert.Contains("background-color:#333333", directTableHtml.QuerySelectorAll("tr")[0].Children[0].GetAttribute("style"));
        Assert.Contains("background-color:transparent", directTableHtml.QuerySelectorAll("tr")[1].Children[0].GetAttribute("style"));

        table._tableProperties.Shading = new Shading { Val = ShadingPatternValues.Nil, Fill = "FF0000" };
        var clearedTableHtml = new HtmlParser().ParseDocument(document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = includeCss }));
        Assert.Contains("background-color:transparent", clearedTableHtml.QuerySelector("table")!.GetAttribute("style"));
        Assert.Contains("background-color:transparent", clearedTableHtml.QuerySelectorAll("tr")[0].Children[0].GetAttribute("style"));
    }

    [Fact]
    public void InheritedBordersFollowMergedCellFramesAndRetainExplicitNilAndNone() {
        using var document = WordDocument.Create();
        var table = document.AddTable(2, 2, WordTableStyle.TableNormal);
        table._tableProperties!.TableBorders = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 8U, Color = "FF0000" },
            new LeftBorder { Val = BorderValues.Single, Size = 24U, Color = "0000FF" },
            new BottomBorder { Val = BorderValues.Single, Size = 16U, Color = "00FF00" },
            new RightBorder { Val = BorderValues.Single, Size = 32U, Color = "000000" },
            new InsideHorizontalBorder { Val = BorderValues.Dotted, Size = 2U, Color = "444444" },
            new InsideVerticalBorder { Val = BorderValues.Dashed, Size = 1U, Color = "666666" });
        table.MergeCells(0, 0, 2, 1);
        table.Rows[1].Cells[1].Borders.TopStyle = WordBorderStyle.Nil;
        table.Rows[1].Cells[1].Borders.LeftStyle = WordBorderStyle.None;

        var html = new HtmlParser().ParseDocument(document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = false }));
        string merged = html.QuerySelector("td[rowspan='2']")!.GetAttribute("style")!;
        string last = html.QuerySelectorAll("tr")[1].Children[0].GetAttribute("style")!;

        Assert.Contains("border-top:1pt solid #FF0000", merged, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("border-bottom:2pt solid #00FF00", merged, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("border-left:3pt solid #0000FF", merged, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("border-right:0.125pt dashed #666666", merged);
        Assert.Contains("border-top:hidden", last);
        Assert.Contains("border-left:none", last);
        Assert.Contains("border-right:4pt solid #000000", last);
    }

    private static string SourceXml(WordDocument document) {
        var main = document.MainDocumentPartRoot;
        return main.Document.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml + main.ThemePart!.Theme.OuterXml +
            main.DocumentSettingsPart!.Settings!.OuterXml;
    }
}

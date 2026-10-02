using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlSemanticTableContractTests {
    [Theory]
    [InlineData("garbage")]
    [InlineData("-1")]
    public void InvalidRowSpanRetainsAdapterLossDiagnostics(string span) {
        var source = HtmlConversionDocument.Parse("<table><tr><td rowspan='" + span + "'>A</td></tr></table>");
        var excelResult = OfficeIMO.Excel.Html.HtmlExcelConverterExtensions.ToExcelDocumentResult(source,
            new OfficeIMO.Excel.Html.HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using var excel = excelResult.RequireValue();
        Assert.Contains(excelResult.Report.Diagnostics, d => d.Code == HtmlConversionDiagnosticCodes.TableSpanInvalid);
        var slideResult = OfficeIMO.PowerPoint.Html.HtmlPowerPointConverterExtensions.ToPowerPointPresentationResult(source,
            new OfficeIMO.PowerPoint.Html.HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, ImportEditableLayoutRegions = false });
        using var slides = slideResult.RequireValue();
        Assert.Contains(slideResult.Report.Diagnostics, d => d.Code == HtmlConversionDiagnosticCodes.TableSpanInvalid);
    }

    [Theory]
    [InlineData("2")]
    [InlineData("99")]
    public void ExplicitRowSpanStopsAtItsRowGroup(string span) {
        var source = HtmlConversionDocument.Parse("<table><tbody><tr><td rowspan='" + span
            + "'>A</td></tr></tbody><tbody><tr><td>B</td></tr></tbody></table>");
        Assert.Equal(1, source.SemanticDocument.RootTables.Single().Table!.Rows[0].Cells[0].RowSpan);
        using var excel = OfficeIMO.Excel.Html.HtmlExcelConverterExtensions.ToExcelDocument(source,
            new OfficeIMO.Excel.Html.HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        Assert.Equal("B", excel.Sheets[0].CellAt(2, 1).GetValue<string>());
    }

    [Fact]
    public void ZeroAriaRowSpanStopsAtItsRowGroupInSemanticAndNativeTargets() {
        const string html = "<div role='table'><div role='rowgroup'><div role='row'><div role='cell' aria-rowspan='0'>Group</div><div role='cell'>A</div></div>"
            + "<div role='row'><div role='cell'>B</div></div></div><div role='rowgroup'><div role='row'><div role='cell'>Next</div></div></div></div>";
        var source = HtmlConversionDocument.Parse(html);
        var table = source.SemanticDocument.RootTables.Single().Table!;
        Assert.Equal(2, table.Rows[0].Cells[0].RowSpan);
        using var excel = OfficeIMO.Excel.Html.HtmlExcelConverterExtensions.ToExcelDocument(source,
            new OfficeIMO.Excel.Html.HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        Assert.Equal("A1:A2", Assert.Single(excel.Sheets[0].GetMergedRanges()).A1Range);
        Assert.Equal("Next", excel.Sheets[0].CellAt(3, 1).GetValue<string>());
        using var slides = OfficeIMO.PowerPoint.Html.HtmlPowerPointConverterExtensions.ToPowerPointPresentation(source,
            new OfficeIMO.PowerPoint.Html.HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, ImportEditableLayoutRegions = false });
        Assert.Equal("Next", Assert.Single(slides.Slides.SelectMany(slide => slide.Tables)).GetCell(2, 0).Text);
    }

    [Fact]
    public void NativeCellsIgnoreAriaSpanWhenNativeAttributeIsAbsent() {
        var table = HtmlConversionDocument.Parse("<table><tr><td aria-rowspan='2'>A</td></tr><tr><td>B</td></tr></table>")
            .SemanticDocument.RootTables.Single().Table!;
        Assert.Equal(1, table.Rows[0].Cells[0].RowSpan);
    }

    [Fact]
    public void RoleTableRetainsHeadersSpansAndSourceRows() {
        var source = HtmlConversionDocument.Parse("<div role='table' aria-label='Values'><div role='row'></div>"
            + "<div role='row'><span role='columnheader' aria-colspan='2'>Name</span></div>"
            + "<div role='rowgroup'><div role='row'><span role='cell'>A</span><span role='cell'>B</span></div></div></div>");
        var table = Assert.Single(source.SemanticDocument.Sections.SelectMany(s => s.Blocks), b => b.Kind == HtmlSemanticBlockKind.Table).Table!;
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(1, table.Rows[0].SourceRowIndex);
        Assert.Equal(2, table.Rows[1].SourceRowIndex);
        Assert.True(table.Rows[0].Cells[0].IsHeader);
        Assert.Equal(2, table.Rows[0].Cells[0].ColumnSpan);
        Assert.Equal(new[] { "A", "B" }, table.Rows[1].Cells.Select(c => c.Text));
    }

    [Fact]
    public void AuthoredCaptionRunsRemainSeparateFromFallbackTitle() {
        var authored = HtmlConversionDocument.Parse("<table><caption>Units <strong>SI</strong></caption><tr><td>m</td></tr></table>")
            .SemanticDocument.Sections.SelectMany(s => s.Blocks).Single(b => b.Table != null).Table!;
        Assert.Equal("Units SI", string.Concat(authored.CaptionRuns.Select(r => r.Text)));
        var fallback = HtmlConversionDocument.Parse("<h2>Units</h2><table><tr><td>m</td></tr></table>")
            .SemanticDocument.Sections.SelectMany(s => s.Blocks).Single(b => b.Table != null).Table!;
        Assert.Empty(fallback.CaptionRuns);
        Assert.Equal("Units", fallback.Caption);
    }

    [Fact]
    public void ImageSemanticHyperlinksUseThePreparedDocumentPolicy() {
        const string image = "<img src='data:image/png;base64,iVBORw0KGgo=' alt='Photo'>";
        var source = HtmlConversionDocument.Parse("<a href='https://example.test/photo'>" + image + "</a>");
        var resource = Assert.Single(source.SemanticDocument.Resources);
        Assert.Equal("https://example.test/photo", resource.Hyperlink);
        var unsafeSource = HtmlConversionDocument.Parse("<a href='javascript:alert(1)'>" + image + "</a>");
        Assert.Null(Assert.Single(unsafeSource.SemanticDocument.Resources).Hyperlink);
    }
}

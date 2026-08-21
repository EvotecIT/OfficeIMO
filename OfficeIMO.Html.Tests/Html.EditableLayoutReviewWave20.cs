using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave20Tests {
    [Theory]
    [InlineData("<div>First</div><div>Second</div>")]
    [InlineData("First<div>Second</div>")]
    [InlineData("<div>First</div>Second")]
    [InlineData("<section><div>First</div><div>Second</div></section>")]
    public void MultipleVisibleBlockContentItemsRemainInSemanticFlow(string content) {
        string html = "<div style='position:absolute;width:180px;height:70px'>" + content + "</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions(),
            HtmlCssMediaContext.Screen);

        Assert.Empty(projection.Regions);
        Assert.Contains("First", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Contains("Second", projection.RemainingDocument.Body.TextContent, StringComparison.Ordinal);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("multipleBlockChildren=true", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("outline:4px solid red")]
    [InlineData("outline-style:solid;outline-width:4px;outline-color:red;outline-offset:2px")]
    public void OutlinedRegionsRemainInSemanticFlow(string outlineStyle) {
        string html = "<div style='position:absolute;width:180px;height:70px;" + outlineStyle + "'>Outlined</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions(),
            HtmlCssMediaContext.Screen);

        Assert.Empty(projection.Regions);
        Assert.Contains("Outlined", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("outline", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("<table></table>")]
    [InlineData("<table></table><table><tbody></tbody></table>")]
    public void EmptyRootTablesDoNotRenumberLaterProjectionWorksheets(string emptyTables) {
        string html = emptyTables
            + "<table><caption>Owned</caption><tr><td>Cell"
            + "<div style='position:absolute;width:140px;height:40px'>Owned region</div>"
            + "</td></tr></table>";

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;
        ExcelSheet sheet = Assert.Single(workbook.Sheets);

        Assert.True(ContainsCellText(sheet, "Owned region"));
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("owning worksheet was not created", StringComparison.Ordinal));
    }

    [Fact]
    public void NonPaintingOutlineDoesNotPreventNativeProjection() {
        const string html = "<div style='position:absolute;width:180px;height:70px;outline:0 solid red'>Plain</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions(),
            HtmlCssMediaContext.Screen);

        Assert.Single(projection.Regions);
        Assert.DoesNotContain(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("outline", StringComparison.OrdinalIgnoreCase));
    }

    private static bool ContainsCellText(ExcelSheet sheet, string expected) {
        string normalizedExpected = string.Concat(expected.Where(character => !char.IsWhiteSpace(character)));
        for (int row = 1; row <= 30; row++) {
            for (int column = 1; column <= 10; column++) {
                if (sheet.TryGetCellText(row, column, out string value)
                    && string.Concat(value.Where(character => !char.IsWhiteSpace(character)))
                        .Contains(normalizedExpected, StringComparison.Ordinal)) return true;
            }
        }
        return false;
    }
}

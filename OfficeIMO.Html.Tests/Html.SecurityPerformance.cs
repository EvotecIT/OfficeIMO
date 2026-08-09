using System;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    private const string SecurityPixel =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII=";

    [Fact]
    public async System.Threading.Tasks.Task HtmlToWord_SharedAncestorStylesAreParsedOncePerElement() {
        const int paragraphCount = 200;
        var ancestorStyle = new StringBuilder("background-color:#0000ff;");
        while (ancestorStyle.Length < 64 * 1024) {
            ancestorStyle.Append("--padding-contract:0;");
        }
        var html = new StringBuilder("<div style=\"")
            .Append(ancestorStyle)
            .Append("\">");
        for (int index = 0; index < paragraphCount; index++) {
            html.Append("<p style=\"background-color:rgba(255,0,0,0.5)\">Text</p>");
        }
        html.Append("</div>");
        HtmlToWordOptions options = HtmlToWordOptions.CreateUntrustedHtmlProfile();
        var converter = new HtmlToWordConverter();

        using WordDocument document = await converter.ConvertAsync(
            HtmlConversionDocument.Parse(html.ToString()).CreateDocumentForConversion(),
            options);

        Assert.Equal(paragraphCount, document.Paragraphs.Count);
        Assert.All(document.Paragraphs, paragraph =>
            Assert.Equal("800080", paragraph.ShadingFillColorHex));
        Assert.Equal(3, converter.InlineStyleParseCount);
    }

    [Fact]
    public async System.Threading.Tasks.Task HtmlToWord_TableBackgroundIsParsedOncePerOwner() {
        const int cellCount = 200;
        var ancestorStyle = new StringBuilder("background-color:#0000ff;");
        while (ancestorStyle.Length < 64 * 1024) {
            ancestorStyle.Append("--table-background-contract:0;");
        }
        var rowBackground = new StringBuilder("rgba(255,0,0,0.5)");
        while (rowBackground.Length < 64 * 1024) {
            rowBackground.Append(' ');
        }
        var html = new StringBuilder("<div style=\"")
            .Append(ancestorStyle)
            .Append("\"><table><tr style=\"background-color:")
            .Append(rowBackground)
            .Append("\">");
        for (int index = 0; index < cellCount; index++) {
            html.Append("<td>Cell</td>");
        }
        html.Append("</tr></table></div>");
        HtmlToWordOptions options = HtmlToWordOptions.CreateUntrustedHtmlProfile();
        var converter = new HtmlToWordConverter();

        using WordDocument document = await converter.ConvertAsync(
            HtmlConversionDocument.Parse(html.ToString()).CreateDocumentForConversion(),
            options);

        WordTable table = Assert.Single(document.Tables);
        Assert.Equal(cellCount, table.Rows[0].Cells.Count);
        Assert.All(table.Rows[0].Cells, cell =>
            Assert.Equal("800080", cell.ShadingFillColorHex));
        Assert.Equal(1, converter.TableBackgroundParseCount);
    }

    [Fact]
    public void HtmlToWord_ImageContainerWidthIsResolvedOnlyForPercentageSizing() {
        using WordDocument document = WordDocument.Create();
        int resolverCalls = 0;
        int? ResolveContainerWidth() {
            resolverCalls++;
            return 7200;
        }

        Assert.Null(HtmlToWordConverter.TryResolveImagePercentWidth(
            "320px",
            document,
            ResolveContainerWidth));
        Assert.Equal(0, resolverCalls);

        double resolvedWidth = Assert.IsType<double>(
            HtmlToWordConverter.TryResolveImagePercentWidth(
                "50%",
                document,
                ResolveContainerWidth));
        Assert.Equal(240D, resolvedWidth, precision: 3);
        Assert.Equal(1, resolverCalls);
    }

    [Fact]
    public async System.Threading.Tasks.Task HtmlToWord_PercentageTableImagesUsePrecomputedCellWidths() {
        string html = $"""
            <table style="width:600px">
              <tr>
                <td style="border:8px solid #000"><img style="width:100%" src="data:image/png;base64,{SecurityPixel}"></td>
                <td style="border:8px solid #000"><img style="width:100%" src="data:image/png;base64,{SecurityPixel}"></td>
                <td style="border:8px solid #000"><img style="width:100%" src="data:image/png;base64,{SecurityPixel}"></td>
              </tr>
            </table>
            """;
        var converter = new HtmlToWordConverter();

        using WordDocument document = await converter.ConvertAsync(
            HtmlConversionDocument.Parse(html).CreateDocumentForConversion(),
            HtmlToWordOptions.CreateUntrustedHtmlProfile());

        Assert.Equal(3, document.Images.Count);
        var field = typeof(HtmlToWordConverter).GetField(
            "_tableCellContentWidths",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        var widths = Assert.IsAssignableFrom<System.Collections.IDictionary>(field!.GetValue(converter));
        Assert.Equal(3, widths.Count);
        Assert.All(widths.Values.Cast<int?>(), width => Assert.True(width > 0));
        WordTable table = Assert.Single(document.Tables);
        var estimateMethod = typeof(WordTable).GetMethod(
            "EstimateCellContentWidthInDxa",
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
        for (int index = 0; index < table.Rows[0].Cells.Count; index++) {
            int expectedWidth = Assert.IsType<int>(estimateMethod!.Invoke(
                null,
                new object[] { document, table.Rows[0].Cells[index]._tableCell }));
            Assert.Equal(expectedWidth / 15D, document.Images[index].Width!.Value, precision: 3);
        }
    }
}

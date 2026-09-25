using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void HtmlToWord_AriaRoleTable_SavesEditableRowsAndCells() {
        const string html = "<div role='table' aria-label='Terms to know'><div role='rowgroup'>" +
            "<div role='row'><div role='columnheader'>Term</div><div role='columnheader'>Definition</div></div>" +
            "<div role='row'><div role='cell'><p><b>Safe water</b></p></div><div role='cell'><p>Available when needed</p></div></div>" +
            "</div></div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument document = result.RequireValue();
        WordTable table = Assert.Single(document.Tables);
        Assert.Equal("Terms to know", table.Title);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(2, table.Rows[0].Cells.Count);
        Assert.Contains(table.Rows[1].Cells[0].Paragraphs, paragraph =>
            paragraph.Text.Contains("Safe water", StringComparison.Ordinal));
        Assert.False(result.Report.HasLoss);

        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);
        Assert.Equal(2, Assert.Single(reopened.Tables).Rows.Count);
        Assert.Contains("Available when needed", reopened.ToHtml(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToWord_AriaRoleTable_RespectsTableCellLimit() {
        const string html = "<div role='table'><div role='row'><div role='cell'>A</div><div role='cell'>B</div></div>" +
            "<div role='row'><div role='cell'>C</div><div role='cell'>D</div></div></div>";

        var exception = Assert.Throws<HtmlConversionLimitException>(() =>
            HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions { MaxTableCells = 3 }));

        Assert.Equal("TableSizeLimitExceeded", exception.Code);
        Assert.Equal(4, exception.Actual);
    }

    [Fact]
    public void HtmlToWord_AriaRoleTable_PreservesOriginalSelectorStyling() {
        const string html = "<style>div[role='table'] > div[role='row'] > div[role='cell'] { background-color:#00ff00; }</style>" +
            "<div role='table'><div role='row'><div role='cell'>Styled cell</div></div></div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument document = result.RequireValue();

        Assert.Equal("00FF00", Assert.Single(document.Tables).Rows[0].Cells[0].ShadingFillColorHex);
    }

    [Fact]
    public void HtmlToWord_AriaRoleTable_ResolvesRemAgainstRootFontBeforeNormalization() {
        const string html = "<style>html { font-size:20px; } div[role='cell'] { font-size:2rem; margin-inline-start:1em; }</style>" +
            "<div role='table'><div role='row'><div role='cell'>Sized cell</div></div></div>";

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            Assert.Single(document.Tables).Rows[0].Cells[0].Paragraphs,
            candidate => candidate.Text == "Sized cell");

        Assert.Equal(600, paragraph.IndentationBefore);
    }

    [Fact]
    public void HtmlToWord_AriaRoleTable_DoesNotApplyRulesForSyntheticNativeTags() {
        const string html = "<style>div[role='cell'] { background-color:#00ff00; } td { background-color:#0000ff !important; }</style>" +
            "<div role='table'><div role='row'><div role='cell'>Original role cell</div></div></div>";

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();

        Assert.Equal("00FF00", Assert.Single(document.Tables).Rows[0].Cells[0].ShadingFillColorHex);
    }

    [Fact]
    public void HtmlToWord_AriaRoleTable_PreservesLinkAndImageCellHosts() {
        string pixel = "data:image/png;base64," +
            Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 3));
        string html = "<div role='table'><div role='row'>" +
            "<a role='cell' href='https://example.org/docs'>Docs</a>" +
            "<img role='cell' src='" + pixel + "' alt='Pixel'>" +
            "</div></div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument document = result.RequireValue();

        Assert.Equal(2, Assert.Single(document.Tables).Rows[0].Cells.Count);
        string roundTrip = document.ToHtml();
        Assert.Contains("https://example.org/docs", roundTrip, StringComparison.Ordinal);
        Assert.Contains("<img", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlToWord_AriaRoleTable_SyntheticNodesDoNotConsumeSourceDepthLimit() {
        const string html = "<div role='table'><div role='row'><div role='cell'></div></div></div>";

        using WordDocument document = HtmlConversionDocument.Parse(html)
            .ToWordDocument(new HtmlToWordOptions { MaxHtmlDepth = 5 });

        Assert.Single(document.Tables);
    }

    [Fact]
    public void HtmlToWord_UnsupportedAriaRoleTable_ReportsApproximation() {
        const string html = "<div role='table'><p>Unstructured content</p><div role='row'><div role='cell'>Value</div></div></div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument document = result.RequireValue();

        Assert.Empty(document.Tables);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated &&
            diagnostic.Source == "role=table");
    }
}

using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void HtmlToWord_TableAccessibilityMetadata_RoundTrips() {
        const string html = "<table aria-label=\"Quarterly results\" aria-description=\"Revenue by region\"><tr><td>North</td></tr></table>";

        using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

        WordTable table = Assert.Single(document.Tables);
        Assert.Equal("Quarterly results", table.Title);
        Assert.Equal("Revenue by region", table.Description);

        string roundTrip = document.ToHtml();
        Assert.Contains("aria-label=\"Quarterly results\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("aria-description=\"Revenue by region\"", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlToWord_TableAccessibilityName_ResolvesAriaLabelledBy() {
        const string html = "<p id=\"table-name\">Regional totals</p><table aria-labelledby=\"table-name\" summary=\"Legacy description\"><tr><td>North</td></tr></table>";

        using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

        WordTable table = Assert.Single(document.Tables);
        Assert.Equal("Regional totals", table.Title);
        Assert.Equal("Legacy description", table.Description);
    }

    [Fact]
    public void HtmlToWord_TableTitleFallbackIsNotDuplicatedAsDescription() {
        const string html = "<table title=\"Quarterly results\"><tr><td>North</td></tr></table>";

        using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

        WordTable table = Assert.Single(document.Tables);
        Assert.Equal("Quarterly results", table.Title);
        Assert.True(string.IsNullOrEmpty(table.Description));
        string roundTrip = document.ToHtml();
        Assert.Contains("aria-label=\"Quarterly results\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("aria-description=\"Quarterly results\"", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlToWord_EmptyDescriptionMetadataFallsBackToTheFirstNonEmptyCarrier() {
        const string html = "<table aria-description=\"  \" summary=\"Legacy description\" title=\"Table title\"><tr><td>North</td></tr></table>";

        using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

        WordTable table = Assert.Single(document.Tables);
        Assert.Equal("Table title", table.Title);
        Assert.Equal("Legacy description", table.Description);
    }

    [Fact]
    public void WordToHtml_ContentControlWrappedNestedTablePreservesBlockOrder() {
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(1, 1);
        WordTable nested = outer.Rows[0].Cells[0].AddTable(1, 1);
        nested.Rows[0].Cells[0].Paragraphs[0].Text = "Wrapped nested";
        nested._table.Remove();
        outer.Rows[0].Cells[0]._tableCell.Append(
            new DocumentFormat.OpenXml.Wordprocessing.SdtBlock(
                new DocumentFormat.OpenXml.Wordprocessing.SdtProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.SdtAlias { Val = "Nested table wrapper" }),
                new DocumentFormat.OpenXml.Wordprocessing.SdtContentBlock(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Before wrapped"))),
                    nested._table,
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("After wrapped"))))));

        string html = document.ToHtml();
        int beforeIndex = html.IndexOf("Before wrapped", StringComparison.Ordinal);
        int nestedIndex = html.IndexOf("Wrapped nested", StringComparison.Ordinal);
        int afterIndex = html.IndexOf("After wrapped", StringComparison.Ordinal);

        Assert.True(beforeIndex >= 0 && nestedIndex > beforeIndex && afterIndex > nestedIndex, html);
        var parsed = HtmlDocumentParser.ParseDocument(html);
        Assert.Single(parsed.QuerySelectorAll("table table"));
    }

    [Fact]
    public void WordToHtml_CustomXmlWrappedNestedTablePreservesBlockOrder() {
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(1, 1);
        WordTable nested = outer.Rows[0].Cells[0].AddTable(1, 1);
        nested.Rows[0].Cells[0].Paragraphs[0].Text = "Custom XML nested";
        nested._table.Remove();
        outer.Rows[0].Cells[0]._tableCell.Append(
            new DocumentFormat.OpenXml.Wordprocessing.CustomXmlBlock(
                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text("Before custom XML"))),
                nested._table,
                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text("After custom XML")))));

        string html = document.ToHtml();
        int beforeIndex = html.IndexOf("Before custom XML", StringComparison.Ordinal);
        int nestedIndex = html.IndexOf("Custom XML nested", StringComparison.Ordinal);
        int afterIndex = html.IndexOf("After custom XML", StringComparison.Ordinal);

        Assert.True(beforeIndex >= 0 && nestedIndex > beforeIndex && afterIndex > nestedIndex, html);
        var parsed = HtmlDocumentParser.ParseDocument(html);
        Assert.Single(parsed.QuerySelectorAll("table table"));
    }

    [Fact]
    public void WordToHtml_NestedTableInsideListItem_PreservesBlockOrder() {
        const string html = "<table><tr><td><ol><li>Before<table><tr><td>Nested</td></tr></table></li><li>After</li></ol></td></tr></table>";

        using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

        string roundTrip = document.ToHtml();
        int beforeIndex = roundTrip.IndexOf("Before", StringComparison.Ordinal);
        int nestedIndex = roundTrip.IndexOf("Nested", StringComparison.Ordinal);
        int afterIndex = roundTrip.IndexOf("After", StringComparison.Ordinal);
        Assert.True(beforeIndex >= 0 && nestedIndex > beforeIndex && afterIndex > nestedIndex, roundTrip);

        var parsed = HtmlDocumentParser.ParseDocument(roundTrip);
        var list = Assert.IsAssignableFrom<AngleSharp.Dom.IElement>(parsed.QuerySelector("td > ol"));
        var items = list.Children.Where(element => element.LocalName == "li").ToArray();
        Assert.Equal(2, items.Length);
        Assert.NotNull(items[0].QuerySelector("table"));
        Assert.Equal("After", items[1].TextContent);
    }

    [Fact]
    public void WordToHtml_SiblingTableAfterClosedListRemainsOutsideListItem() {
        const string html = "<table><tr><td><ol><li>List item</li></ol><table><tr><td>Sibling table</td></tr></table></td></tr></table>";

        using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
        var parsed = HtmlDocumentParser.ParseDocument(document.ToHtml());

        Assert.Empty(parsed.QuerySelectorAll("li table"));
        Assert.Single(parsed.QuerySelectorAll("td > table"));
    }

    [Fact]
    public void WordToHtml_NestedTableAsLastListItemChildUsesExactRoundTripMarker() {
        const string html = "<table><tr><td><ol><li>List item<table><tr><td>Nested table</td></tr></table></li></ol></td></tr></table>";

        using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
        var parsed = HtmlDocumentParser.ParseDocument(document.ToHtml());

        Assert.Single(parsed.QuerySelectorAll("li > table"));
    }

    [Fact]
    public void WordToHtml_OuterListTableMarkerDoesNotLeakToDescendantSiblingTable() {
        const string html = "<table><tr><td><ol><li>Outer<table><tr><td><ol><li>Inner</li></ol><table><tr><td>Inner sibling table</td></tr></table></td></tr></table></li></ol></td></tr></table>";

        using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
        var parsed = HtmlDocumentParser.ParseDocument(document.ToHtml());

        var listTables = parsed.QuerySelectorAll("li > table");
        Assert.Single(listTables);
        Assert.Contains("Inner sibling table", listTables[0].TextContent, StringComparison.Ordinal);
        Assert.Empty(parsed.QuerySelectorAll("li li > table"));
    }
}

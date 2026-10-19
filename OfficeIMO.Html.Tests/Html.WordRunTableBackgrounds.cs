using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Tests;

public sealed class HtmlWordRunTableBackgroundTests {
    [Theory]
    [InlineData("whole", false)]
    [InlineData("first-row", false)]
    [InlineData("inherited", false)]
    [InlineData("default", false)]
    [InlineData("nil", false)]
    [InlineData("clear", false)]
    [InlineData("first-row-nil", false)]
    [InlineData("whole", true)]
    [InlineData("first-row", true)]
    [InlineData("none", true)]
    [InlineData("first-row-none", true)]
    public void TableRunBackgroundOverridesDocumentDefaultsInHtmlAfterReopen(string source, bool highlight) {
        using WordDocument document = CreateTableBackground(source, highlight);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string html = loaded.ToHtml(new WordToHtmlOptions { IncludeRunHighlightStyles = true });
        IElement text = HtmlDocumentParser.ParseDocument(html).QuerySelectorAll("td *, th *")
            .Single(element => element.TextContent == "BACKGROUND" && element.Children.Length == 0);

        Assert.Equal(source is "nil" or "clear" or "none" or "first-row-nil" or "first-row-none" ? null : "#0000ff", ReadInlineBackground(text));
    }

    [Theory]
    [InlineData("semantic", false)]
    [InlineData("semantic", true)]
    [InlineData("equation", false)]
    [InlineData("equation", true)]
    [InlineData("marked", false)]
    [InlineData("marked", true)]
    [InlineData("marked-equation", false)]
    [InlineData("marked-equation", true)]
    public void SemanticAndEquationAdjacentTextUsesTableHighlightAndExplicitNoneAfterReopen(string wrapper, bool reset) {
        using WordDocument document = CreateTableBackground(reset ? "none" : "first-row", highlight: true);
        Paragraph paragraph = document.Tables[0]._table.Descendants<Paragraph>().Single();
        foreach (Run run in paragraph.Elements<Run>().ToArray()) run.Remove();
        var text = new Run(new Text("BACKGROUND"));
        if (wrapper.StartsWith("marked", StringComparison.Ordinal))
            text.RunProperties = new RunProperties(new RunStyle { Val = "HtmlMarkedText" });
        if (wrapper is "semantic" or "marked") {
            if (wrapper == "semantic") text.RunProperties = new RunProperties(new RunStyle { Val = "HtmlCode" });
            paragraph.Append(text);
        } else {
            paragraph.Append(new Hyperlink(text, new M.OfficeMath(new M.Run(new M.Text("x")))) { Anchor = "equation-target" });
        }

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        var html = HtmlDocumentParser.ParseDocument(loaded.ToHtml(new WordToHtmlOptions { IncludeRunHighlightStyles = true }));
        IElement output = html.QuerySelectorAll("td *, th *")
            .Single(element => element.ChildNodes.Any(node => node.NodeType == NodeType.Text && node.TextContent == "BACKGROUND"));
        Assert.Equal(reset ? (wrapper.StartsWith("marked", StringComparison.Ordinal) ? "transparent" : null) : "#0000ff", ReadInlineBackground(output));
        if (wrapper == "marked") Assert.Equal("BACKGROUND", html.QuerySelector("td mark, th mark")?.TextContent);
        else if (wrapper == "semantic") Assert.Equal("BACKGROUND", html.QuerySelector("td code, th code")?.TextContent);
        else {
            IElement anchor = Assert.IsAssignableFrom<IElement>(html.QuerySelector("a[href='#equation-target']"));
            Assert.NotNull(anchor.QuerySelector("math"));
            Assert.Contains("BACKGROUND", anchor.TextContent, StringComparison.Ordinal);
        }
    }


    [Fact]
    public void RepeatedConversionsRefreshTablePositionsAndConditionalPaint() {
        using WordDocument document = CreateTableBackground("first-row", highlight: true);
        WordTable table = document.Tables[0];
        table.AddRow(1).Cells[0].Paragraphs[0].Text = "SECOND";
        string before = document.ToHtml();
        Assert.Equal("#0000ff", BackgroundFor(before, "BACKGROUND"));
        Assert.Equal("#ffff00", BackgroundFor(before, "SECOND"));
        TableRow first = table._table.Elements<TableRow>().First();
        first.Remove(); table._table.Append(first);
        string after = document.ToHtml();
        Assert.Equal("#ffff00", BackgroundFor(after, "BACKGROUND"));
        Assert.Equal("#0000ff", BackgroundFor(after, "SECOND"));
    }

    private static string? BackgroundFor(string html, string label) => ReadInlineBackground(
        HtmlDocumentParser.ParseDocument(html).QuerySelectorAll("td *, th *")
            .Single(element => element.ChildNodes.Any(node => node.NodeType == NodeType.Text && node.TextContent == label)));

    private static string? ReadInlineBackground(IElement text) {
        for (IElement? element = text; element != null; element = element.ParentElement) {
            string? declaration = element.GetAttribute("style")?.Split(';')
                .LastOrDefault(part => part.TrimStart().StartsWith("background-color:", StringComparison.OrdinalIgnoreCase));
            if (declaration != null) return declaration.Substring(declaration.IndexOf(':') + 1).Trim().ToLowerInvariant();
            if (element.LocalName is "td" or "th") break;
        }
        return null;
    }

    private static WordDocument CreateTableBackground(string source, bool highlight) {
        WordDocument document = WordDocument.Create();
        Styles styles = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(highlight
            ? new RunPropertiesBaseStyle(new Highlight { Val = HighlightColorValues.Yellow })
            : new RunPropertiesBaseStyle(Shade("FF0000"))));
        var properties = new RunPropertiesBaseStyle();
        if (highlight) properties.Append(new Highlight { Val = source.EndsWith("none", StringComparison.Ordinal) ? HighlightColorValues.None : HighlightColorValues.Blue });
        else properties.Append(source.EndsWith("nil", StringComparison.Ordinal) ? new Shading { Val = ShadingPatternValues.Nil, Fill = "0000FF" }
            : Shade(source == "clear" ? "auto" : "0000FF"));
        var style = new Style { Type = StyleValues.Table, StyleId = "TableBackground" };
        bool conditional = source.StartsWith("first-row", StringComparison.Ordinal) || source == "inherited";
        if (conditional)
            style.Append(new TableStyleProperties(properties) { Type = TableStyleOverrideValues.FirstRow });
        else style.StyleRunProperties = new StyleRunProperties(properties.ChildElements.Select(child => child.CloneNode(true)));
        styles.Append(style);
        string styleId = "TableBackground";
        if (source == "inherited") {
            styles.Append(new Style(new BasedOn { Val = styleId }, new StyleRunProperties(Shade("00FF00"))) {
                Type = StyleValues.Table, StyleId = "DerivedTableBackground"
            });
            styleId = "DerivedTableBackground";
        }
        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
        table._tableProperties.TableLook = new TableLook { Val = "0600", FirstRow = conditional };
        if (source == "default") {
            foreach (Style other in styles.Elements<Style>().Where(candidate => candidate.Type?.Value == StyleValues.Table)) other.Default = false;
            style.Default = true;
            table._tableProperties.TableStyle = null;
        }
        table.Rows[0].Cells[0].Paragraphs[0].Text = "BACKGROUND";
        return document;
    }

    private static Shading Shade(string fill) => new Shading { Val = ShadingPatternValues.Clear, Fill = fill };
}

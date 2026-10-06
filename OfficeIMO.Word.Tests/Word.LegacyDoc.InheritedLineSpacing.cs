using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body")]
    [InlineData("clear-points")]
    [InlineData("list")]
    [InlineData("table")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("footnote")]
    [InlineData("endnote")]
    [InlineData("comment")]
    [InlineData("derived-style")]
    [InlineData("table-style")]
    [InlineData("table-conditional")]
    [InlineData("table-conditional-base")]
    [InlineData("table-value")]
    [InlineData("root-style")]
    [InlineData("table-conditional-explicit")]
    public void LegacyDoc_LineRuleWithoutValueUsesTheInheritedSpacingPair(string route) {
        using WordDocument source = WordDocument.Create();
        var mainPart = source._wordprocessingDocument.MainDocumentPart!;
        Styles styles = mainPart.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults!.ParagraphPropertiesDefault!.ParagraphPropertiesBaseStyle!.SpacingBetweenLines =
            new SpacingBetweenLines { Line = "440", LineRule = LineSpacingRuleValues.Exact };
        source.AddParagraph("Anchor");
        WordParagraph target;
        if (route == "header" || route == "footer") {
            source.AddHeadersAndFooters();
            target = route == "header" ? source.Header.Default!.AddParagraph("Inherited") : source.Footer.Default!.AddParagraph("Inherited");
        } else if (route == "footnote") {
            target = source.AddParagraph("Note").AddFootNote("Inherited").FootNote!.Paragraphs!.Last();
        } else if (route == "endnote") {
            target = source.AddParagraph("Note").AddEndNote("Inherited").EndNote!.Paragraphs!.Last();
        } else if (route == "comment") {
            source.AddParagraph("Commented").AddComment("OfficeIMO", "OI", "Inherited");
            target = source.Comments.Single().Paragraphs.Single();
        } else if (route == "list") {
            var list = source.AddCustomList();
            list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
            target = list.AddItem("Inherited");
        } else if (route.StartsWith("table", StringComparison.Ordinal)) {
            var table = source.AddTable(1, 1);
            table._tableProperties!.TableStyle = null;
            target = table.Rows[0].Cells[0].Paragraphs[0];
            target.Text = "Inherited";
            if (route != "table") {
                var properties = new StyleParagraphProperties(new SpacingBetweenLines {
                    Line = route == "table-value" ? "600" : route == "table-conditional-explicit" ? "420" : null,
                    LineRule = route == "table-value" ? LineSpacingRuleValues.Exact : LineSpacingRuleValues.Auto
                });
                var tableStyle = new Style(new StyleName { Val = "Inherited table spacing" }) {
                    StyleId = "InheritedTableSpacing", Type = StyleValues.Table, CustomStyle = true
                };
                if (route is "table-style" or "table-value") tableStyle.Append(properties);
                else tableStyle.Append(new TableStyleProperties(properties) { Type = TableStyleOverrideValues.FirstRow });
                if (route is "table-conditional-base" or "table-conditional-explicit") {
                    styles.Append(new Style(new StyleName { Val = "Base table spacing" },
                        new TableStyleProperties(new StyleParagraphProperties(new SpacingBetweenLines {
                            Line = "360", LineRule = LineSpacingRuleValues.Exact
                        })) { Type = TableStyleOverrideValues.FirstRow }) {
                        Type = StyleValues.Table, StyleId = "BaseTableSpacing", CustomStyle = true
                    });
                    tableStyle.BasedOn = new BasedOn { Val = "BaseTableSpacing" };
                    tableStyle.StyleParagraphProperties = new StyleParagraphProperties(new SpacingBetweenLines {
                        Line = "440", LineRule = LineSpacingRuleValues.Exact
                    });
                }
                styles.Append(tableStyle);
                table._tableProperties.TableStyle = new TableStyle { Val = "InheritedTableSpacing" };
                table.ConditionalFormattingFirstRow = true;
            }
        } else target = source.AddParagraph("Inherited");
        foreach (Style style in styles.Elements<Style>().Where(style => style.Type?.Value == StyleValues.Paragraph)) {
            if (style.StyleParagraphProperties != null) style.StyleParagraphProperties.SpacingBetweenLines = null;
        }
        if (route == "derived-style") {
            styles.Append(new Style(new StyleName { Val = "Inherited rule" }, new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new SpacingBetweenLines { LineRule = LineSpacingRuleValues.Auto })) {
                StyleId = "InheritedRule", Type = StyleValues.Paragraph, CustomStyle = true
            });
            target.SetStyleId("InheritedRule");
        } else if (route == "root-style") {
            Style normal = styles.Elements<Style>().Single(style => style.StyleId?.Value == "Normal");
            normal.StyleParagraphProperties ??= new StyleParagraphProperties();
            normal.StyleParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines { LineRule = LineSpacingRuleValues.Auto };
        } else if (!route.StartsWith("table-", StringComparison.Ordinal) || route == "table-value") {
            target.LineSpacingRule = route == "clear-points" ? WordLineSpacingRule.AtLeast : WordLineSpacingRule.Auto;
            if (route == "clear-points") {
                target.LineSpacingPoints = 18; target.LineSpacingPoints = null;
                Assert.Null(target.LineSpacing);
                Assert.Equal(WordLineSpacingRule.AtLeast, target.LineSpacingRule);
                Assert.Null(target._paragraph.ParagraphProperties!.SpacingBetweenLines!.Line);
            }
        }
        target.FontSize = 32;
        var originalRoots = new[] { mainPart.Document }.Cast<OpenXmlPartRootElement>()
            .Concat(mainPart.Parts.Select(part => part.OpenXmlPart.RootElement).OfType<OpenXmlPartRootElement>())
            .Select(root => (Root: root, Xml: root.OuterXml)).ToArray();
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        Assert.All(originalRoots, item => Assert.Equal(item.Xml, item.Root.OuterXml));
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        var reopenedMainPart = reopened._wordprocessingDocument.MainDocumentPart!;
        var reopenedRoots = new[] { reopenedMainPart.Document }.Cast<OpenXmlPartRootElement>()
            .Concat(reopenedMainPart.Parts.Select(part => part.OpenXmlPart.RootElement).OfType<OpenXmlPartRootElement>());
        Paragraph paragraph = Assert.Single(reopenedRoots.SelectMany(root => root.Descendants<Paragraph>()),
            paragraph => paragraph.InnerText == "Inherited");
        SpacingBetweenLines? spacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
        string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        while (string.IsNullOrWhiteSpace(spacing?.Line?.Value) && styleId != null && visited.Add(styleId)) {
            Style? style = reopenedMainPart.StyleDefinitionsPart!.Styles!.Elements<Style>().SingleOrDefault(style => style.StyleId?.Value == styleId);
            spacing = style?.StyleParagraphProperties?.SpacingBetweenLines;
            styleId = style?.BasedOn?.Val?.Value;
        }
        Assert.True(spacing != null, paragraph.OuterXml);
        Assert.Equal(route == "table-value" ? "600" : route == "table-conditional-base" ? "360" : route == "table-conditional-explicit" ? "420" : "440", spacing!.Line?.Value);
        Assert.Equal(route == "table-conditional-explicit" ? LineSpacingRuleValues.Auto : LineSpacingRuleValues.Exact, spacing.LineRule?.Value);
    }
}

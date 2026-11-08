using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlThemeApplicationTests {
    private static OfficeTheme CustomTheme() => OfficeTheme.Modern
        .WithFonts(new OfficeThemeFonts(new OfficeThemeFont("Georgia", Array.Empty<string>()),
            new OfficeThemeFont("Arial", Array.Empty<string>())))
        .WithTypeScale(OfficeTheme.Modern.TypeScale.With(body: OfficeLength.Points(20), heading1: OfficeLength.Points(31), heading2: OfficeLength.Points(27)))
        .WithTextStyle(OfficeTheme.Modern.TextStyle.With(headingColor: OfficeColor.Parse("2468AC"), headingBold: false))
        .WithTableStyle(OfficeTheme.Modern.TableStyle.With(headerFill: OfficeColor.Parse("112233"),
            headerText: OfficeColor.Parse("FFFFFF"), headerBold: false, cellPadding: OfficeLength.Points(4)))
        .WithPage(new OfficeThemePage(OfficeLength.Points(18), OfficeLength.Points(24), OfficeLength.Points(30), OfficeLength.Points(36)));

    [Fact]
    public void HtmlThemeDefaultsPreserveAuthoredCascadeAndIndependentSnapshots() {
        const string input = "<style>:where(h1){color:#00aa00;font-size:16pt}</style>" +
            "<h1>Authored heading</h1><h2>Theme heading</h2><p><strong>Meaning</strong> <a href='details'>Link</a></p>" +
            "<table><thead><tr><th>Declared header</th></tr></thead><tbody><tr><td>Body</td></tr></tbody></table>" +
            "<table><tr><td id='plain'>Ordinary first row</td></tr></table>";
        var source = HtmlConversionDocument.Parse(input, new HtmlConversionDocumentOptions {
            BaseUri = new Uri("https://example.test/reports/input.html"),
            Limits = new HtmlConversionLimits { MaxInputCharacters = 20000, MaxHtmlNodes = 100 }
        });
        var themed = source.WithTheme(CustomTheme());
        var styles = HtmlComputedStyleEngine.Compute(themed);
        Assert.Equal(16D, styles[themed.Document.QuerySelector("h1")!].ResolvedFontSizePoints);
        Assert.Equal("rgba(0, 170, 0, 1)", styles[themed.Document.QuerySelector("h1")!].GetValue("color"));
        Assert.Equal(27D, styles[themed.Document.QuerySelector("h2")!].ResolvedFontSizePoints);
        Assert.Equal("rgba(36, 104, 172, 1)", styles[themed.Document.QuerySelector("h2")!].GetValue("color"));
        Assert.Equal(20D, styles[themed.Document.QuerySelector("p")!].ResolvedFontSizePoints);
        Assert.Contains("Georgia", styles[themed.Document.QuerySelector("h2")!].GetValue("font-family"));
        Assert.Equal("rgba(17, 34, 51, 1)", styles[themed.Document.QuerySelector("th")!].GetValue("background-color"));
        Assert.NotEqual("rgba(17, 34, 51, 1)", styles[themed.Document.QuerySelector("#plain")!].GetValue("background-color"));
        Assert.Equal(source.Document.Body!.TextContent, themed.Document.Body!.TextContent);
        Assert.Equal("details", themed.Document.QuerySelector("a")!.GetAttribute("href"));
        Assert.Equal(source.BaseUri, themed.BaseUri);
        Assert.Equal(source.Trust, themed.Trust);
        Assert.Equal(100, themed.Limits.MaxHtmlNodes);
        Assert.Equal(input, source.SourceHtml);
        Assert.Empty(source.Document.QuerySelectorAll("style[data-officeimo-theme]"));

        var switched = themed.WithTheme(CustomTheme().WithTypeScale(CustomTheme().TypeScale.With(heading2: OfficeLength.Points(33))));
        Assert.Single(switched.Document.QuerySelectorAll("style[data-officeimo-theme]"));
        Assert.Equal(33D, HtmlComputedStyleEngine.Compute(switched)[switched.Document.QuerySelector("h2")!].ResolvedFontSizePoints);
        Assert.Equal(27D, HtmlComputedStyleEngine.Compute(themed)[themed.Document.QuerySelector("h2")!].ResolvedFontSizePoints);
        var relocated = HtmlConversionDocument.Parse(switched.ExportSourceHtml());
        Assert.Equal(switched.BaseUri, relocated.BaseUri);
        Assert.Equal(33D, HtmlComputedStyleEngine.Compute(relocated)[relocated.Document.QuerySelector("h2")!].ResolvedFontSizePoints);

        var empty = HtmlConversionDocument.Parse("").Edit(document => document.DocumentElement!.Remove());
        Assert.NotNull(empty.WithTheme(CustomTheme()).Document.Head);
        Assert.Null(empty.Document.DocumentElement);
    }

    [Fact]
    public void HtmlWordThemeKeepsAdapterFontAuthoredCssAndHeaderSemantics() {
        var html = HtmlConversionDocument.Parse("<style>@page{margin-left:25pt}#authored{font-family:'Courier New';font-size:13pt}</style>" +
            "<p id='authored'>Authored face</p><h2>Default heading</h2>" +
            "<table><thead><tr><th>Declared header</th></tr></thead><tbody><tr><td>Data</td></tr></tbody></table>" +
            "<table><tr><td>Ordinary first row</td></tr></table>");
        using WordDocument document = html.ToWord(new HtmlWordExportOptions { Theme = CustomTheme(), FontFamily = "Verdana" });
        byte[] bytes = document.ToBytes();
        using var stream = new MemoryStream(bytes, false);
        using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Paragraph authored = package.MainDocumentPart!.Document.Descendants<Paragraph>().Single(p => p.InnerText == "Authored face");
        Paragraph heading = package.MainDocumentPart.Document.Descendants<Paragraph>().Single(p => p.InnerText == "Default heading");
        Assert.Equal("Courier New", authored.Descendants<RunFonts>().First().Ascii!.Value);
        Assert.Equal("26", authored.Descendants<FontSize>().First().Val!.Value);
        Style headingStyle = package.MainDocumentPart.StyleDefinitionsPart!.Styles!.Elements<Style>()
            .Single(style => style.StyleId == heading.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        string? headingSize = heading.Descendants<FontSize>().FirstOrDefault()?.Val?.Value ?? headingStyle.StyleRunProperties?.FontSize?.Val?.Value;
        PageMargin margins = package.MainDocumentPart.Document.Descendants<PageMargin>().First();
        Table[] tables = package.MainDocumentPart.Document.Descendants<Table>().ToArray();
        TableHeader? declared = tables[0].Elements<TableRow>().First().TableRowProperties?.GetFirstChild<TableHeader>();
        TableHeader? ordinary = tables[1].Elements<TableRow>().First().TableRowProperties?.GetFirstChild<TableHeader>();
        Assert.Equal(("Verdana", "54", 360, 500U, true, false),
            (heading.Descendants<RunFonts>().First().Ascii!.Value, headingSize, margins.Top!.Value, margins.Left!.Value,
                declared != null && declared.Val?.Value != OnOffOnlyValues.Off, ordinary != null && ordinary.Val?.Value != OnOffOnlyValues.Off));
        Assert.Empty(html.Document.QuerySelectorAll("style[data-officeimo-theme]"));

        var snapshot = HtmlConversionDocument.Parse("<style>*{color:#556677}th{border:none;border-bottom:1pt solid #112233;" +
            "border-bottom-width:2pt;border-bottom-color:#A12345}</style>" +
            "<h2>Snapshot heading</h2><p>Relative body <strong>Semantic strong</strong> <em>Semantic emphasis</em> " +
            "<strong style='font-weight:initial'>Reset weight</strong> <em style='font-style:initial'>Reset slant</em> " +
            "<u style='text-decoration:none'>Reset underline</u> <s style='text-decoration:none'>Reset strike</s> " +
            "<var style='font-style:normal'>Reset variable</var></p>" +
            "<p><span style='display:grid;background-image:linear-gradient(red,blue)'>Unsupported authored effects</span></p>" +
            "<table><thead><tr><th>Snapshot header</th></tr></thead><tbody><tr><td>Snapshot data</td></tr></tbody></table>")
            .WithTheme(CustomTheme());
        // A serialized snapshot carries its own theme; adapter Theme remains omitted.
        snapshot = HtmlConversionDocument.Parse(snapshot.ExportSourceHtml());
        var snapshotOptions = new HtmlWordExportOptions();
        snapshotOptions.StylesheetContents.Add("@layer report{h2{font-size:17pt}@page{margin-left:19pt}}" +
            "@media screen{p{font-size:150%}}@media print{p{font-size:9pt}}");
        using WordDocument snapshotWord = snapshot.ToWord(snapshotOptions, out OfficeConversionReport snapshotReport);
        using var snapshotStream = new MemoryStream(snapshotWord.ToBytes(), false);
        using WordprocessingDocument snapshotPackage = WordprocessingDocument.Open(snapshotStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(snapshotPackage));
        Paragraph snapshotHeading = Assert.Single(snapshotPackage.MainDocumentPart!.Document.Descendants<Paragraph>(), paragraph => paragraph.InnerText == "Snapshot heading");
        Run snapshotHeadingRun = Assert.Single(snapshotHeading.Descendants<Run>(), run => run.InnerText == "Snapshot heading");
        Assert.Equal(("Georgia", "34", "556677"),
            (snapshotHeadingRun.RunProperties!.RunFonts!.Ascii!.Value, snapshotHeadingRun.RunProperties.FontSize!.Val!.Value,
                snapshotHeadingRun.RunProperties.Color!.Val!.Value));
        Assert.False(snapshotHeadingRun.RunProperties.Bold!.Val!.Value);
        Assert.Equal(380U, snapshotPackage.MainDocumentPart.Document.Descendants<PageMargin>().First().Left!.Value);
        Run relative = Assert.Single(snapshotPackage.MainDocumentPart.Document.Descendants<Run>(), run => run.InnerText.Trim() == "Relative body");
        Assert.Equal("60", relative.RunProperties!.FontSize!.Val!.Value);
        Run reset = Assert.Single(snapshotPackage.MainDocumentPart.Document.Descendants<Run>(), run => run.InnerText == "Reset weight");
        Assert.False(reset.RunProperties!.Bold!.Val!.Value);
        Assert.True(Assert.Single(snapshotPackage.MainDocumentPart.Document.Descendants<Run>(), run => run.InnerText == "Semantic strong")
            .RunProperties!.Bold!.Val!.Value);
        Assert.True(Assert.Single(snapshotPackage.MainDocumentPart.Document.Descendants<Run>(), run => run.InnerText == "Semantic emphasis")
            .RunProperties!.Italic!.Val!.Value);
        Assert.False(Assert.Single(snapshotPackage.MainDocumentPart.Document.Descendants<Run>(), run => run.InnerText == "Reset slant")
            .RunProperties!.Italic!.Val!.Value);
        Assert.Equal(UnderlineValues.None, Assert.Single(snapshotPackage.MainDocumentPart.Document.Descendants<Run>(), run => run.InnerText == "Reset underline")
            .RunProperties!.Underline!.Val!.Value);
        Assert.False(Assert.Single(snapshotPackage.MainDocumentPart.Document.Descendants<Run>(), run => run.InnerText == "Reset strike")
            .RunProperties!.Strike!.Val!.Value);
        Assert.False(Assert.Single(snapshotPackage.MainDocumentPart.Document.Descendants<Run>(), run => run.InnerText == "Reset variable")
            .RunProperties!.Italic!.Val!.Value);
        TableCell snapshotHeader = Assert.Single(snapshotPackage.MainDocumentPart.Document.Descendants<TableCell>(), cell => cell.InnerText == "Snapshot header");
        Assert.Equal("112233", snapshotHeader.TableCellProperties!.Shading!.Fill!.Value);
        Assert.Equal("288", snapshotHeader.GetFirstChild<Paragraph>()!.ParagraphProperties!.SpacingBetweenLines!.Line!.Value);
        Assert.Equal(BorderValues.None, snapshotHeader.TableCellProperties.TableCellBorders!.TopBorder!.Val!.Value);
        Assert.Equal((BorderValues.Single, 16U, "A12345"),
            (snapshotHeader.TableCellProperties.TableCellBorders!.BottomBorder!.Val!.Value,
                snapshotHeader.TableCellProperties.TableCellBorders.BottomBorder.Size!.Value,
                snapshotHeader.TableCellProperties.TableCellBorders.BottomBorder.Color!.Value));
        Assert.DoesNotContain(snapshotReport.Diagnostics, diagnostic => diagnostic.Code.StartsWith("unsupported-css-", StringComparison.Ordinal)
            && (diagnostic.Location?.EndsWith(":line-height", StringComparison.Ordinal) == true ||
                diagnostic.Location?.EndsWith(":border-collapse", StringComparison.Ordinal) == true ||
                diagnostic.Location?.StartsWith("th:background", StringComparison.Ordinal) == true ||
                diagnostic.Location?.StartsWith("th:border-", StringComparison.Ordinal) == true));
        Assert.Contains(snapshotReport.Diagnostics, diagnostic => diagnostic.Code == "unsupported-css-declaration" && diagnostic.Location == "table:margin-bottom");
        Assert.Contains(snapshotReport.Diagnostics, diagnostic => diagnostic.Code == "unsupported-css-declaration" && diagnostic.Location == "span:background-image");
        Assert.Contains(snapshotReport.Diagnostics, diagnostic => diagnostic.Code == "unsupported-css-declaration" && diagnostic.Location == "span:display");
        Assert.Null(snapshotOptions.Theme);

        using WordDocument authoredLayer = HtmlConversionDocument.Parse("<style>@layer report{body{background-color:#123456;color:#2468ac}" +
            "#layered{font-weight:700}}p{color:#556677}</style><p id='layered'>Authored layer</p>").ToWord();
        using var layeredStream = new MemoryStream(authoredLayer.ToBytes(), false);
        using WordprocessingDocument layeredPackage = WordprocessingDocument.Open(layeredStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(layeredPackage));
        Assert.Equal("123456", layeredPackage.MainDocumentPart!.Document.DocumentBackground!.Color!.Value);
        Run layeredRun = Assert.Single(layeredPackage.MainDocumentPart.Document.Descendants<Run>(), run => run.InnerText == "Authored layer");
        Assert.Equal("556677", layeredRun.RunProperties!.Color!.Val!.Value);
        Assert.True(layeredRun.RunProperties.Bold!.Val!.Value);

        var pageOptions = new HtmlWordExportOptions { DefaultPageSize = WordPageSize.Letter, DefaultOrientation = OfficePageOrientation.Portrait };
        pageOptions.StylesheetContents.Add("@page{margin:9pt}");
        var paged = HtmlConversionDocument.Parse("<style>@page{margin-left:25pt;size:A5 landscape}@page:first{margin-top:2pt}</style>" +
            "<style media='screen'>@page{margin:3pt}</style><section class='word-section' data-word-section='1' data-margin-bottom-twips='700'><p>Page geometry</p></section>");
        using WordDocument configured = paged.ToWord(pageOptions, out OfficeConversionReport pageReport);
        using var configuredStream = new MemoryStream(configured.ToBytes(), false);
        using WordprocessingDocument configuredPackage = WordprocessingDocument.Open(configuredStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(configuredPackage));
        SectionProperties configuredSection = configuredPackage.MainDocumentPart!.Document.Descendants<SectionProperties>().Single();
        PageMargin configuredMargins = configuredSection.GetFirstChild<PageMargin>()!;
        PageSize configuredSize = configuredSection.GetFirstChild<PageSize>()!;
        Assert.Equal((180, 500U, 180U, 700),
            (configuredMargins.Top!.Value, configuredMargins.Left!.Value, configuredMargins.Right!.Value, configuredMargins.Bottom!.Value));
        Assert.Equal((11906U, 8391U, PageOrientationValues.Landscape),
            (configuredSize.Width!.Value, configuredSize.Height!.Value, configuredSize.Orient!.Value));
        Assert.Contains(pageReport.Diagnostics, diagnostic => diagnostic.Code == "html-page-specific-rules-omitted");

        using WordDocument fallback = HtmlConversionDocument.Parse("<p>Native defaults</p>").ToWord(new HtmlWordExportOptions {
            DefaultPageSize = WordPageSize.Letter, DefaultOrientation = OfficePageOrientation.Landscape
        });
        Assert.Equal((15840L, 12240L, 1440L),
            (fallback.PageSetup.Size.Width!.Value.ToTwips(), fallback.PageSetup.Size.Height!.Value.ToTwips(), fallback.PageSetup.Margins.Top!.Value.ToTwips()));
        string sectionBeforeInsertion = fallback.Sections.Single()._sectionProperties.OuterXml;
        fallback.AddHtml("<style>@page{margin:1pt;size:A5}</style><p>Inserted content</p>", new HtmlWordExportOptions { Theme = CustomTheme() });
        Assert.Equal(sectionBeforeInsertion, fallback.Sections.Single()._sectionProperties.OuterXml);
    }
}

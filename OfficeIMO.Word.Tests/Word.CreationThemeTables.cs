using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public class WordCreationThemeTables {
    private static readonly XNamespace WordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private sealed class ThemeRow {
        public string Name { get; set; } = "Alpha";
        public int Count { get; set; } = 12;
    }

    private static WordTable AddRows(WordDocument document, string route, OfficeTableOptions? options = null) => route switch {
        "typed" => document.AddTable(new[] { new ThemeRow() }, options),
        "matrix" => document.AddTable(new object?[,] { { "Name", "Count" }, { "Alpha", 12 } }, options),
        "text" => document.AddTable(new IReadOnlyList<string>[] { new[] { "Name", "Count" }, new[] { "Alpha", "12" } }, options),
        _ => throw new ArgumentOutOfRangeException(nameof(route))
    };

    private static OfficeTheme CustomTheme(string accent, string bodyFont) {
        OfficeTheme baseline = OfficeTheme.Modern;
        return baseline.WithName("Custom " + accent)
            .WithPalette(baseline.Palette.With(accent1: OfficeColor.ParseHex(accent), text: OfficeColor.ParseHex("17324D")))
            .WithFonts(new OfficeThemeFonts(new OfficeThemeFont("Georgia", new[] { "serif" }), new OfficeThemeFont(bodyFont, new[] { "sans-serif" })))
            .WithTableStyle(baseline.TableStyle.With(headerFill: OfficeColor.Accent1, headerText: OfficeColor.Background,
                bandFill: OfficeColor.ParseHex("E7F0EE"), cellPadding: OfficeLength.Points(7), numberAlignment: OfficeAlignment.Center));
    }

    private static (XDocument Styles, XDocument Theme, XDocument Body) ReadNative(byte[] bytes) {
        using var input = new MemoryStream(bytes);
        using var package = WordprocessingDocument.Open(input, false);
        return (XDocument.Parse(package.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml),
            XDocument.Parse(package.MainDocumentPart.ThemePart!.Theme!.OuterXml),
            XDocument.Parse(package.MainDocumentPart.Document.OuterXml));
    }

    private static XElement ModernStyle(XDocument styles) => styles.Root!.Elements(WordNs + "style")
        .Single(style => (string?)style.Attribute(WordNs + "styleId") == "OfficeIMOModern");

    private static XElement HeaderStyle(XElement style) => style.Elements(WordNs + "tblStylePr")
        .Single(element => (string?)element.Attribute(WordNs + "type") == "firstRow");

    private static void AssertNativeTheme(byte[] bytes, OfficeTheme theme) {
        var native = ReadNative(bytes);
        XElement style = native.Styles.Root!.Elements(WordNs + "style").Single(element => (string?)element.Attribute(WordNs + "styleId") == "OfficeIMOModern");
        XElement header = HeaderStyle(style);
        Assert.Equal(theme.TableStyle.HeaderFill.Resolve(theme).ToRgbHex(),
            (string?)header.Descendants(WordNs + "shd").Single().Attribute(WordNs + "fill"));
        Assert.Equal(theme.TableStyle.HeaderText.Resolve(theme).ToRgbHex(),
            (string?)header.Descendants(WordNs + "color").Single().Attribute(WordNs + "val"));
        Assert.Equal(theme.Fonts.Body.Family, (string?)style.Element(WordNs + "rPr")!.Element(WordNs + "rFonts")!.Attribute(WordNs + "ascii"));
        Assert.Equal(theme.Palette.Accent1.ToRgbHex(), (string?)native.Theme.Descendants(DrawingNs + "accent1")
            .Single().Element(DrawingNs + "srgbClr")!.Attribute("val"));
        Assert.Equal(theme.Palette.Text.ToRgbHex(), (string?)native.Theme.Descendants(DrawingNs + "dk1")
            .Single().Element(DrawingNs + "srgbClr")!.Attribute("val"));
        Assert.Equal(theme.Fonts.Heading.Family, (string?)native.Theme.Descendants(DrawingNs + "majorFont")
            .Single().Element(DrawingNs + "latin")!.Attribute("typeface"));
        Assert.Equal(theme.Fonts.Body.Family, (string?)native.Theme.Descendants(DrawingNs + "minorFont")
            .Single().Element(DrawingNs + "latin")!.Attribute("typeface"));
        Assert.Equal(theme.Fonts.Body.Family, (string?)native.Styles.Descendants(WordNs + "docDefaults")
            .Single().Descendants(WordNs + "rFonts").Single().Attribute(WordNs + "ascii"));
        string padding = theme.TableStyle.CellPadding.ToTwips().ToString(CultureInfo.InvariantCulture);
        Assert.Equal(padding, (string?)style.Descendants(WordNs + "tblCellMar").Single().Element(WordNs + "top")!.Attribute(WordNs + "w"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextColorRolesPreserveNativeThemeSlotsAndLiteralOverrides(bool loaded) {
        byte[] seed;
        using (var source = WordDocument.Create(new WordCreateOptions { Theme = CustomTheme("13579B", "Tahoma") }))
            seed = source.ToBytes();
        using var input = new MemoryStream(seed);
        using var document = loaded ? WordDocument.Load(input) :
            WordDocument.Create(new WordCreateOptions { Theme = CustomTheme("13579B", "Tahoma") });
        var text = document.AddParagraph("Accent text");
        text.Color = OfficeColor.Accent1;
        var link = document.AddParagraph().AddHyperLink("Accent link", new Uri("https://example.org"));
        link.Runs[0].Color = OfficeColor.Accent1;
        var native = ReadNative(document.ToBytes());
        foreach (string label in new[] { "Accent text", "Accent link" }) {
            var color = native.Body.Descendants(WordNs + "r")
                .Single(run => run.Element(WordNs + "t")?.Value == label)
                .Element(WordNs + "rPr")!.Element(WordNs + "color")!;
            Assert.Equal("13579B", (string?)color.Attribute(WordNs + "val"));
            Assert.Equal("accent1", (string?)color.Attribute(WordNs + "themeColor"));
        }
        text.Color = OfficeColor.ParseHex("2468AC");
        var literal = ReadNative(document.ToBytes()).Body.Descendants(WordNs + "r")
            .Single(run => run.Element(WordNs + "t")?.Value == "Accent text")
            .Element(WordNs + "rPr")!.Element(WordNs + "color")!;
        Assert.Equal("2468AC", (string?)literal.Attribute(WordNs + "val"));
        Assert.Null(literal.Attribute(WordNs + "themeColor"));
    }

    [Theory]
    [InlineData("typed", false)]
    [InlineData("matrix", false)]
    [InlineData("text", false)]
    [InlineData("typed", true)]
    [InlineData("matrix", true)]
    [InlineData("text", true)]
    public void AllRoutesUseNativeModernStyleAndSelectedTheme(string route, bool custom) {
        OfficeTheme theme = custom ? CustomTheme("157A67", "Arial") : OfficeTheme.Modern;
        using var document = WordDocument.Create(custom ? new WordCreateOptions { Theme = theme } : null);
        WordTable table = AddRows(document, route);
        Assert.Equal(WordTableStyle.OfficeIMOModern, table.Style);
        Assert.True(table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage);
        Assert.Equal(custom ? WordParagraphAlignment.Center : WordParagraphAlignment.Right, table.Rows[1].Cells[1].Paragraphs[0].ParagraphAlignment);
        byte[] bytes = document.ToBytes();
        AssertNativeTheme(bytes, theme);
        var native = ReadNative(bytes);
        Assert.Equal("OfficeIMOModern", (string?)native.Body.Descendants(WordNs + "tblStyle").Single().Attribute(WordNs + "val"));
        if (!custom) {
            XElement style = native.Styles.Root!.Elements(WordNs + "style").Single(element => (string?)element.Attribute(WordNs + "styleId") == "OfficeIMOModern");
            Assert.Equal("EAF0FC", (string?)HeaderStyle(style).Descendants(WordNs + "shd").Single().Attribute(WordNs + "fill"));
            Assert.Equal("20", (string?)style.Element(WordNs + "rPr")!.Element(WordNs + "sz")!.Attribute(WordNs + "val"));
            Assert.Equal("F6F8FB", (string?)style.Elements(WordNs + "tblStylePr")
                .Single(element => (string?)element.Attribute(WordNs + "type") == "band2Horz")
                .Descendants(WordNs + "shd").Single().Attribute(WordNs + "fill"));
        }
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).Where(error =>
            error.Part?.Uri.OriginalString == "/word/styles.xml" || error.Part?.Uri.OriginalString == "/word/theme/theme1.xml"));
    }

    [Theory]
    [InlineData("typed")]
    [InlineData("matrix")]
    [InlineData("text")]
    public void HeaderlessRoutesDisableHeaderTreatment(string route) {
        using var document = WordDocument.Create();
        WordTable table = AddRows(document, route, new OfficeTableOptions { IncludeHeaders = false });
        Assert.False(table.ConditionalFormattingFirstRow);
        Assert.False(table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage);
        var native = ReadNative(document.ToBytes());
        XElement look = native.Body.Descendants(WordNs + "tblLook").Single();
        string? expanded = (string?)look.Attribute(WordNs + "firstRow");
        bool firstRow = expanded != null ? expanded is "1" or "true" or "on"
            : (int.Parse((string?)look.Attribute(WordNs + "val") ?? "0", NumberStyles.HexNumber, CultureInfo.InvariantCulture) & 0x0020) != 0;
        Assert.False(firstRow);
    }

    [Fact]
    public void ProjectionFormattingAndExplicitAlignmentWinAndSurviveReload() {
        OfficeTheme theme = CustomTheme("157A67", "Arial");
        var options = new OfficeTableOptions();
        options.Order.Add("Count");
        options.HeaderRenames["Count"] = "Quantity";
        options.Formatters["Count"] = value => "Count=" + value;
        options.Alignments.Add(OfficeAlignment.Left);
        using var document = WordDocument.Create(new WordCreateOptions { Theme = theme });
        WordTable table = AddRows(document, "typed", options);
        WordParagraph value = table.Rows[1].Cells[0].Paragraphs[0];
        Assert.Equal("Quantity", table.Rows[0].Cells[0].Paragraphs[0].Text);
        Assert.Equal("Count=12", value.Text);
        Assert.Equal(WordParagraphAlignment.Left, value.ParagraphAlignment);
        value.Color = OfficeColor.ParseHex("AF1234");
        value.Bold = false;
        value.FontSize = OfficeLength.Points(17);
        byte[] bytes = document.ToBytes();
        AssertNativeTheme(bytes, theme);
        using var loaded = WordDocument.Load(new MemoryStream(bytes));
        WordParagraph reopened = loaded.Tables[0].Rows[1].Cells[0].Paragraphs[0];
        Assert.Equal(OfficeColor.ParseHex("AF1234"), reopened.Color);
        Assert.False(reopened.Bold);
        Assert.Equal(OfficeLength.Points(17), reopened.FontSize);
        Assert.Equal(WordParagraphAlignment.Left, reopened.ParagraphAlignment);
        Assert.Equal(ModernStyle(ReadNative(bytes).Styles).ToString(), ModernStyle(ReadNative(loaded.ToBytes()).Styles).ToString());
    }

    [Theory]
    [InlineData("typed")]
    [InlineData("matrix")]
    [InlineData("text")]
    public void ExplicitNumericAlignmentOverridesTheSharedDefault(string route) {
        using var document = WordDocument.Create(new WordCreateOptions { Theme = CustomTheme("157A67", "Arial") });
        var options = new OfficeTableOptions();
        options.Alignments.Add(null);
        options.Alignments.Add(OfficeAlignment.Left);
        WordTable table = AddRows(document, route, options);
        Assert.Equal(WordParagraphAlignment.Left, table.Rows[1].Cells[1].Paragraphs[0].ParagraphAlignment);
        Assert.Equal(WordParagraphAlignment.Left, table.Rows[0].Cells[1].Paragraphs[0].ParagraphAlignment);
        WordParagraph value = table.Rows[1].Cells[1].Paragraphs[0];
        value.Color = OfficeColor.ParseHex("AF1234");
        value.Text = "13";
        Assert.Equal(WordParagraphAlignment.Left, value.ParagraphAlignment);
        Assert.Equal(OfficeColor.ParseHex("AF1234"), value.Color);
    }

    [Fact]
    public async Task ConcurrentDocumentsAndSubsequentDefaultHaveIndependentNativeStyles() {
        OfficeTheme[] themes = { CustomTheme("157A67", "Arial"), CustomTheme("A14673", "Verdana") };
        Task<byte[]>[] tasks = Enumerable.Range(0, 12).Select(index => Task.Run(() => {
            OfficeTheme theme = themes[index % themes.Length];
            using var document = WordDocument.Create(new WordCreateOptions { Theme = theme });
            AddRows(document, "typed");
            return document.ToBytes();
        })).ToArray();
        byte[][] results = await Task.WhenAll(tasks);
        for (int index = 0; index < results.Length; index++) AssertNativeTheme(results[index], themes[index % themes.Length]);
        using var ordinary = WordDocument.Create();
        AddRows(ordinary, "matrix");
        AssertNativeTheme(ordinary.ToBytes(), OfficeTheme.Modern);
    }

    [Theory]
    [InlineData("detached")]
    [InlineData("path")]
    [InlineData("stream")]
    public void CreationRoutesApplyChosenThemeWithoutEagerDestinationWrites(string route) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        using var stream = new MemoryStream();
        OfficeTheme theme = CustomTheme("157A67", "Arial");
        var options = new WordCreateOptions { Theme = theme };
        try {
            using WordDocument document = route switch {
                "path" => WordDocument.Create(path, options),
                "stream" => WordDocument.Create(stream, options),
                _ => WordDocument.Create(options)
            };
            AddRows(document, "matrix");
            Assert.False(File.Exists(path));
            Assert.Equal(0L, stream.Length);
            AssertNativeTheme(document.ToBytes(), theme);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData("detached")]
    [InlineData("path")]
    [InlineData("stream")]
    public async Task TemplatePrecedenceAndOrdinaryLoadPreserveAuthoredNativeContent(string route) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        string destination = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        using var stream = new MemoryStream();
        OfficeTheme original = CustomTheme("157A67", "Arial");
        OfficeTheme selected = CustomTheme("A14673", "Verdana");
        try {
            using (var template = WordDocument.Create(new WordCreateOptions { Theme = original })) {
                AddRows(template, "typed");
                // Native authored style content proves preservation independently of the theme adapter.
                W.Style heading = template.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.Elements<W.Style>()
                    .First(style => style.StyleId?.Value == "Heading1");
                heading.StyleRunProperties ??= new W.StyleRunProperties();
                heading.StyleRunProperties.Color = new W.Color { Val = "AB1278" };
                template.Save(path);
            }
            byte[] authored = File.ReadAllBytes(path);
            var before = ReadNative(authored);
            WordDocument Create(OfficeTheme? theme) {
                var options = new WordCreateOptions { Template = path, Theme = theme };
                return route switch {
                    "path" => WordDocument.Create(destination, options),
                    "stream" => WordDocument.Create(stream, options),
                    _ => WordDocument.Create(options)
                };
            }
            using (var copy = Create(null)) {
                var native = ReadNative(copy.ToBytes());
                Assert.Equal(before.Styles.ToString(), native.Styles.ToString());
                Assert.Equal(before.Theme.ToString(), native.Theme.ToString());
            }
            using (var copy = Create(selected)) {
                byte[] bytes = copy.ToBytes();
                AssertNativeTheme(bytes, selected);
                var after = ReadNative(bytes);
                XElement Heading(XDocument styles) => styles.Root!.Elements(WordNs + "style")
                    .Single(style => (string?)style.Attribute(WordNs + "styleId") == "Heading1");
                Assert.Equal(Heading(before.Styles).ToString(), Heading(after.Styles).ToString());
            }
            using (var loaded = WordDocument.Load(new MemoryStream(authored))) {
                var native = ReadNative(loaded.ToBytes());
                Assert.Equal(before.Styles.ToString(), native.Styles.ToString());
                Assert.Equal(before.Theme.ToString(), native.Theme.ToString());
            }
            using (var loaded = await WordDocument.LoadAsync(new MemoryStream(authored))) {
                var native = ReadNative(loaded.ToBytes());
                Assert.Equal(before.Styles.ToString(), native.Styles.ToString());
                Assert.Equal(before.Theme.ToString(), native.Theme.ToString());
            }
            Assert.Equal(authored, File.ReadAllBytes(path));
            Assert.False(File.Exists(destination));
            Assert.Equal(0L, stream.Length);
        } finally {
            if (File.Exists(path)) File.Delete(path);
            if (File.Exists(destination)) File.Delete(destination);
        }
    }

    [Theory]
    [InlineData(OfficeThemeTableBorderMode.Horizontal)]
    [InlineData(OfficeThemeTableBorderMode.Grid)]
    [InlineData(OfficeThemeTableBorderMode.None)]
    public void CustomBorderModeAndBandingReachNativeStyle(OfficeThemeTableBorderMode mode) {
        OfficeTheme theme = OfficeTheme.Modern.WithTableStyle(OfficeTheme.Modern.TableStyle.With(
            borderMode: mode, banding: false, headerBold: false, headerRule: OfficeThemeRule.None));
        using var document = WordDocument.Create(new WordCreateOptions { Theme = theme });
        AddRows(document, "typed");
        XElement style = ModernStyle(ReadNative(document.ToBytes()).Styles);
        XElement borders = style.Element(WordNs + "tblPr")!.Element(WordNs + "tblBorders")!;
        Assert.Equal(mode == OfficeThemeTableBorderMode.None ? "nil" : "single", (string?)borders.Element(WordNs + "bottom")!.Attribute(WordNs + "val"));
        Assert.Equal(mode == OfficeThemeTableBorderMode.Grid, borders.Element(WordNs + "insideV") != null);
        Assert.DoesNotContain(style.Elements(WordNs + "tblStylePr"), element => (string?)element.Attribute(WordNs + "type") == "band2Horz");
        Assert.Contains((string?)HeaderStyle(style).Descendants(WordNs + "b").Single().Attribute(WordNs + "val"), new[] { "0", "false" });
    }

    [Fact]
    public void ArrayRowsProduceCellsRatherThanArrayMetadata() {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(new[] { new[] { "Name", "Count" }, new[] { "Ada", "12" } });
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(2, table.Rows[0].Cells.Count);
        var native = ReadNative(document.ToBytes());
        Assert.Equal(new[] { "Name", "Count", "Ada", "12" }, native.Body.Descendants(WordNs + "tbl").Single()
            .Descendants(WordNs + "tc").Select(cell => string.Concat(cell.Descendants(WordNs + "t").Select(text => text.Value))));
    }

    [Fact]
    public void ExplicitHeaderBoldOffWritesNativeOverridesForBothScriptKinds() {
        using var document = WordDocument.Create();
        WordTable table = AddRows(document, "typed");
        table.Rows[0].Cells[0].Paragraphs[0].Bold = false;
        var native = ReadNative(document.ToBytes());
        XElement run = native.Body.Descendants(WordNs + "tbl").First().Elements(WordNs + "tr").First()
            .Elements(WordNs + "tc").First().Descendants(WordNs + "r").First();
        foreach (string flag in new[] { "b", "bCs" })
            Assert.Contains((string?)run.Element(WordNs + "rPr")!.Element(WordNs + flag)!.Attribute(WordNs + "val"), new[] { "0", "false" });
        XElement style = native.Styles.Root!.Elements(WordNs + "style").Single(element => (string?)element.Attribute(WordNs + "styleId") == "OfficeIMOModern");
        Assert.Contains((string?)style.Elements(WordNs + "tblStylePr").Single(element =>
            (string?)element.Attribute(WordNs + "type") == "firstRow").Element(WordNs + "rPr")!.Element(WordNs + "b")!.Attribute(WordNs + "val"), new[] { "1", "true" });
    }

    [Fact]
    public void DarkThemeKeepsReadableTextOnWordLightPagesAndBodyBands() {
        using var document = WordDocument.Create(new WordCreateOptions { Theme = OfficeTheme.Dark });
        document.AddParagraph("Page text");
        AddRows(document, "typed");
        var native = ReadNative(document.ToBytes());
        string pageText = OfficeTheme.Dark.Palette.Background.ToRgbHex();
        Assert.Equal(pageText, (string?)native.Styles.Descendants(WordNs + "rPrDefault").Single().Element(WordNs + "rPr")!
            .Element(WordNs + "color")!.Attribute(WordNs + "val"));
        XElement style = native.Styles.Root!.Elements(WordNs + "style").Single(element => (string?)element.Attribute(WordNs + "styleId") == "OfficeIMOModern");
        Assert.Equal(pageText, (string?)style.Element(WordNs + "rPr")!.Element(WordNs + "color")!.Attribute(WordNs + "val"));
        XElement Band(string kind) => style.Elements(WordNs + "tblStylePr").Single(element => (string?)element.Attribute(WordNs + "type") == kind);
        Assert.Equal(OfficeTheme.Modern.TableStyle.BandFill.Resolve(OfficeTheme.Modern).ToRgbHex(),
            (string?)Band("band2Horz").Element(WordNs + "tcPr")!.Element(WordNs + "shd")!.Attribute(WordNs + "fill"));
        Assert.Equal(OfficeTheme.Dark.TableStyle.HeaderFill.Resolve(OfficeTheme.Dark).ToRgbHex(),
            (string?)Band("firstRow").Element(WordNs + "tcPr")!.Element(WordNs + "shd")!.Attribute(WordNs + "fill"));
    }

    [Fact]
    public void ChoosingAThemePreservesExistingHeadingCatalogAndParagraphSpacing() {
        using var baseline = WordDocument.Create();
        using var custom = WordDocument.Create(new WordCreateOptions { Theme = CustomTheme("157A67", "Arial") });
        var before = ReadNative(baseline.ToBytes());
        var after = ReadNative(custom.ToBytes());
        XElement Heading(XDocument styles) => styles.Root!.Elements(WordNs + "style")
            .Single(style => (string?)style.Attribute(WordNs + "styleId") == "Heading1");
        Assert.Equal(Heading(before.Styles).ToString(), Heading(after.Styles).ToString());
        Assert.Equal(before.Styles.Descendants(WordNs + "pPrDefault").Single().ToString(),
            after.Styles.Descendants(WordNs + "pPrDefault").Single().ToString());
        WordParagraph paragraph = custom.AddParagraph("Precise spacing");
        paragraph.LineSpacing = WordLineSpacing.Exact(OfficeLength.Points(13.5));
        using var reopened = WordDocument.Load(new MemoryStream(custom.ToBytes()));
        Assert.Equal(WordLineSpacing.Exact(OfficeLength.Points(13.5)), reopened.Paragraphs[0].LineSpacing);
    }

    [Fact]
    public void DefaultGeometryFillsTheCurrentSectionAndExplicitGridIsStillAvailable() {
        using var document = WordDocument.Create();
        document.Sections.Last().PageSettings.Width = OfficeLength.Inches(10);
        document.Sections.Last().Margins.Left = OfficeLength.Inches(1);
        document.Sections.Last().Margins.Right = OfficeLength.Inches(1);
        WordTable table = AddRows(document, "typed");
        Assert.Equal(WordTableStyle.OfficeIMOModern, table.Style);
        document.AddTable(1, 2, WordTableStyle.TableGrid);
        var native = ReadNative(document.ToBytes());
        XElement first = native.Body.Descendants(WordNs + "tbl").First();
        Assert.Equal(new[] { "5760", "5760" }, first.Element(WordNs + "tblGrid")!.Elements(WordNs + "gridCol")
            .Select(column => (string?)column.Attribute(WordNs + "w")).ToArray());
        Assert.Equal("TableGrid", (string?)native.Body.Descendants(WordNs + "tbl").Last()
            .Element(WordNs + "tblPr")!.Element(WordNs + "tblStyle")!.Attribute(WordNs + "val"));
    }
}

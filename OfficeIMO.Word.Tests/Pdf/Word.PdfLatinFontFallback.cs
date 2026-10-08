using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfLatinFontFallbackTests {
    [Theory]
    [InlineData("body", false)]
    [InlineData("header", false)]
    [InlineData("footer", false)]
    [InlineData("body", true)]
    [InlineData("header", true)]
    [InlineData("footer", true)]
    public void DocumentDefaultsTryHighAnsiBeforeApplicationFont(string story, bool themed) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Courier New";
        var main = document._wordprocessingDocument.MainDocumentPart!;
        main.ThemePart!.Theme!.ThemeElements!.FontScheme!.MajorFont!.LatinFont!.Typeface = "OfficeIMO Unavailable Font";
        main.ThemePart.Theme.ThemeElements.FontScheme.MinorFont!.LatinFont!.Typeface = "Times New Roman";
        main.StyleDefinitionsPart!.Styles!.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.RunFonts = new RunFonts {
            Ascii = "OfficeIMO Unavailable Font", HighAnsi = themed ? "Arial" : "Times New Roman",
            AsciiTheme = themed ? ThemeFontValues.MajorAscii : null,
            HighAnsiTheme = themed ? ThemeFontValues.MinorHighAnsi : null
        };
        AddStoryText(document, story);
        string original = main.Document.OuterXml + main.StyleDefinitionsPart.Styles.OuterXml;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Fallback");
        Assert.All(word.Letters, letter => Assert.Contains("Times", letter.FontName, StringComparison.OrdinalIgnoreCase));
        Assert.Equal(original, main.Document.OuterXml + main.StyleDefinitionsPart.Styles.OuterXml);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    public void SuppliedNonstandardHighAnsiFontIsRegisteredAfterUnavailableAscii(string story) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = AddStoryText(document, story);
        paragraph._run!.RunProperties = new RunProperties(new RunFonts {
            Ascii = "OfficeIMO Unavailable Font", HighAnsi = ManagedTextShapingTestAssets.FamilyName
        });
        var options = new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = options,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Fallback");
        Assert.All(word.Letters, letter => Assert.Contains("OfficeIMO", letter.FontName, StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("character")]
    [InlineData("table")]
    [InlineData("conditional")]
    public void StyleHighAnsiFallbackRetainsSuppliedFont(string layer) {
        using WordDocument document = WordDocument.Create();
        var fonts = new RunFonts { Ascii = "OfficeIMO Unavailable Font", HighAnsi = ManagedTextShapingTestAssets.FamilyName };
        var style = new Style(new StyleName { Val = "Fallback style" }) {
            StyleId = "FallbackStyle", Type = layer is "table" or "conditional" ? StyleValues.Table :
                layer == "character" ? StyleValues.Character : StyleValues.Paragraph
        };
        if (layer == "conditional") style.Append(new TableStyleProperties(new RunPropertiesBaseStyle(fonts)) { Type = TableStyleOverrideValues.FirstRow });
        else style.StyleRunProperties = new StyleRunProperties(fonts);
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
        WordParagraph paragraph;
        if (layer is "table" or "conditional") {
            WordTable table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle = new TableStyle { Val = "FallbackStyle" };
            table.ConditionalFormattingFirstRow = true;
            paragraph = table.Rows[0].Cells[0].Paragraphs[0]; paragraph.Text = "Fallback";
        } else {
            paragraph = document.AddParagraph("Fallback");
            if (layer == "character") paragraph._run!.RunProperties = new RunProperties(new RunStyle { Val = "FallbackStyle" });
            else paragraph.SetStyleId("FallbackStyle");
        }
        var options = new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = options,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.Contains("OfficeIMO", letter.FontName, StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false, "Times")]
    [InlineData(true, "Courier")]
    public void RegistrationRetainsFirstUsableSourceAndFinalDocumentFallback(bool unavailable, string expected) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Courier New";
        WordParagraph paragraph = document.AddParagraph("Fallback");
        paragraph._run!.RunProperties = new RunProperties(new RunFonts {
            Ascii = unavailable ? "OfficeIMO Unavailable Font" : "Times New Roman",
            HighAnsi = unavailable ? "OfficeIMO Other Unavailable Font" : ManagedTextShapingTestAssets.FamilyName
        });
        var options = new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = options, FontFamily = "Courier New",
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.Contains(expected, letter.FontName, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void BlankParagraphHighAnsiFallbackUsesSuppliedLineMetrics() {
        using WordDocument document = WordDocument.Create();
        var first = document.AddParagraph("A");
        var blank = document.AddParagraph();
        var last = document.AddParagraph("B");
        foreach (WordParagraph paragraph in new[] { first, blank, last }) {
            paragraph.LineSpacingPoints = 20; paragraph.LineSpacingRule = WordLineSpacingRule.AtLeast;
            paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
            foreach (Run run in paragraph._paragraph.Descendants<Run>())
                (run.RunProperties ??= new RunProperties()).FontSize = new FontSize { Val = "16" };
        }
        blank._paragraph.RemoveAllChildren<Run>();
        blank._paragraph.ParagraphProperties!.ParagraphMarkRunProperties = new ParagraphMarkRunProperties(
            new RunFonts { Ascii = "OfficeIMO Unavailable Font", HighAnsi = ManagedTextShapingTestAssets.FamilyName },
            new FontSize { Val = "64" });
        var options = new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(1200, -300, 0, 300, Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = options, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(68D, Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.Y -
            Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData("paragraph", false)]
    [InlineData("character", false)]
    [InlineData("table", false)]
    [InlineData("conditional", false)]
    [InlineData("cell", false)]
    [InlineData("paragraph", true)]
    [InlineData("character", true)]
    [InlineData("table", true)]
    [InlineData("conditional", true)]
    [InlineData("cell", true)]
    public void PartialStyleDeclarationsInheritTheOtherLatinSlot(string layer, bool baseAscii) {
        using WordDocument document = WordDocument.Create();
        var inherited = new RunFonts { Ascii = baseAscii ? "Times New Roman" : null,
            HighAnsi = baseAscii ? null : ManagedTextShapingTestAssets.FamilyName };
        var declared = new RunFonts { Ascii = baseAscii ? null : "OfficeIMO Unavailable Font",
            HighAnsi = baseAscii ? ManagedTextShapingTestAssets.FamilyName : null };
        var type = layer is "table" or "conditional" or "cell" ? StyleValues.Table :
            layer == "character" ? StyleValues.Character : StyleValues.Paragraph;
        var parent = new Style(new StyleName { Val = "Parent" }) { StyleId = "FallbackParent", Type = type };
        var child = new Style(new StyleName { Val = "Child" }, new BasedOn { Val = "FallbackParent" }) {
            StyleId = "FallbackChild", Type = type
        };
        if (layer == "conditional") {
            parent.Append(new TableStyleProperties(new RunPropertiesBaseStyle(inherited)) { Type = TableStyleOverrideValues.FirstRow });
            child.Append(new TableStyleProperties(new RunPropertiesBaseStyle(declared)) { Type = TableStyleOverrideValues.FirstRow });
        } else {
            parent.StyleRunProperties = new StyleRunProperties(inherited);
            if (layer == "cell") child.Append(new TableStyleProperties(new RunPropertiesBaseStyle(declared)) { Type = TableStyleOverrideValues.FirstRow });
            else child.StyleRunProperties = new StyleRunProperties(declared);
        }
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(parent, child);
        if (type == StyleValues.Table) {
            WordTable table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle = new TableStyle { Val = "FallbackChild" };
            table.ConditionalFormattingFirstRow = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Fallback";
        } else {
            WordParagraph paragraph = document.AddParagraph("Fallback");
            if (type == StyleValues.Character) paragraph._run!.RunProperties = new RunProperties(new RunStyle { Val = "FallbackChild" });
            else paragraph.SetStyleId("FallbackChild");
        }
        var options = new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = options, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.Contains(baseAscii ? "Times" : "OfficeIMO", letter.FontName, StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false, 11D)]
    [InlineData(true, 22D)]
    public void ExplicitPdfDefaultControlsVisibleAndBlankParagraphLineMetrics(bool blank, double expectedAdvance) {
        using WordDocument document = WordDocument.Create();
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!
            .RunPropertiesDefault!.RunPropertiesBaseStyle!.RunFonts = new RunFonts {
                Ascii = "OfficeIMO Unavailable Font", HighAnsi = "Times New Roman"
            };
        document.AddParagraph("A");
        if (blank) document.AddParagraph();
        document.AddParagraph("B");
        foreach (WordParagraph paragraph in document.Paragraphs) {
            paragraph.LineSpacing = 240; paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
            paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
        }
        var options = new PdfOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica,
            ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(800, -200, 0, 200, Enumerable.Range(32, 95).ToArray()),
            ManagedTextShapingTestAssets.FamilyName);
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = options, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var letters = pdf.GetPage(1).Letters;
        Assert.All(letters, letter => Assert.Contains("OfficeIMO", letter.FontName, StringComparison.Ordinal));
        Assert.Equal(expectedAdvance, Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.Y -
            Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
    }

    private static WordParagraph AddStoryText(WordDocument document, string story) {
        if (story == "body") return document.AddParagraph("Fallback");
        document.AddParagraph("Body"); document.AddHeadersAndFooters();
        return (story == "footer" ? (WordHeaderFooter)document.Footer.Default : document.Header.Default).AddParagraph("Fallback");
    }
}

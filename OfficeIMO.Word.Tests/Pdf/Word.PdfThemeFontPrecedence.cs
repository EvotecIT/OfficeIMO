using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordPdfThemeFontPrecedenceTests {
    [Theory]
    [InlineData("direct")]
    [InlineData("paragraph")]
    [InlineData("character")]
    [InlineData("table")]
    [InlineData("conditional")]
    [InlineData("defaults")]
    public void ThemeSelectorsOverrideLiteralFallbacksInPdf(string layer) {
        using WordDocument document = WordDocument.Create();
        var main = document._wordprocessingDocument.MainDocumentPart!;
        main.ThemePart!.Theme!.ThemeElements!.FontScheme!.MajorFont!.LatinFont!.Typeface = "Times New Roman";
        RunFonts fonts = new() { Ascii = "Arial", HighAnsi = "Arial", AsciiTheme = ThemeFontValues.MajorAscii, HighAnsiTheme = ThemeFontValues.MajorHighAnsi };
        Styles styles = main.StyleDefinitionsPart!.Styles!;
        WordParagraph paragraph;
        if (layer is "table" or "conditional") {
            Style style = new(new StyleName { Val = "Themed table" }) { Type = StyleValues.Table, StyleId = "ThemedTable" };
            if (layer == "conditional") style.Append(new TableStyleProperties(new RunPropertiesBaseStyle(fonts)) { Type = TableStyleOverrideValues.FirstRow });
            else style.StyleRunProperties = new StyleRunProperties(fonts);
            styles.Append(style);
            WordTable table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle = new TableStyle { Val = "ThemedTable" };
            table.ConditionalFormattingFirstRow = true;
            paragraph = table.Rows[0].Cells[0].Paragraphs[0]; paragraph.Text = "Theme font";
        } else {
            paragraph = document.AddParagraph("Theme font");
            if (layer == "direct") paragraph._run!.RunProperties = new RunProperties(fonts);
            else if (layer == "defaults") styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.RunFonts = fonts;
            else {
                styles.Append(new Style(new StyleName { Val = "Themed style" }, new StyleRunProperties(fonts)) {
                    Type = layer == "character" ? StyleValues.Character : StyleValues.Paragraph, StyleId = "ThemedStyle"
                });
                if (layer == "character") paragraph._run!.RunProperties = new RunProperties(new RunStyle { Val = "ThemedStyle" });
                else paragraph.SetStyleId("ThemedStyle");
            }
        }
        string before = main.Document.OuterXml + styles.OuterXml;
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal("Themefont", string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value)).Replace(" ",""));
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.Contains("Times", letter.FontName, StringComparison.OrdinalIgnoreCase));
        Assert.Equal(before, main.Document.OuterXml + styles.OuterXml);
    }

    [Theory]
    [InlineData(false, false, "Times")]
    [InlineData(true, false, "Helvetica")]
    [InlineData(false, true, "Helvetica")]
    public void FontSelectionPreservesSlotOrderAndLiteralFallback(bool primaryLiteral, bool missingTheme, string expected) {
        using WordDocument document = WordDocument.Create();
        var main = document._wordprocessingDocument.MainDocumentPart!;
        main.ThemePart!.Theme!.ThemeElements!.FontScheme!.MajorFont!.LatinFont!.Typeface = "Times New Roman";
        if (missingTheme) main.DeletePart(main.ThemePart);
        WordParagraph paragraph = document.AddParagraph("Font slot");
        paragraph._run!.RunProperties = new RunProperties(new RunFonts {
            Ascii = primaryLiteral ? "Arial" : null, HighAnsi = "Arial", HighAnsiTheme = ThemeFontValues.MajorHighAnsi
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.Contains(expected, letter.FontName, StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("body", false)]
    [InlineData("header", false)]
    [InlineData("footer", false)]
    [InlineData("body", true)]
    [InlineData("header", true)]
    [InlineData("footer", true)]
    public void UnavailableAsciiFamilyRetainsUsableHighAnsiFallback(string story, bool themed) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Courier New";
        var main = document._wordprocessingDocument.MainDocumentPart!;
        main.ThemePart!.Theme!.ThemeElements!.FontScheme!.MajorFont!.LatinFont!.Typeface = "OfficeIMO Unavailable Font";
        main.ThemePart.Theme.ThemeElements.FontScheme.MinorFont!.LatinFont!.Typeface = "Times New Roman";
        WordParagraph paragraph;
        if (story == "body") paragraph = document.AddParagraph("Fallback");
        else {
            document.AddParagraph("Body"); document.AddHeadersAndFooters();
            paragraph = (story == "footer" ? (WordHeaderFooter)document.Footer.Default : document.Header.Default).AddParagraph("Fallback");
        }
        paragraph._run!.RunProperties = new RunProperties(new RunFonts {
            Ascii = "OfficeIMO Unavailable Font", HighAnsi = "Times New Roman",
            AsciiTheme = themed ? ThemeFontValues.MajorAscii : null,
            HighAnsiTheme = themed ? ThemeFontValues.MinorHighAnsi : null
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Fallback");
        Assert.All(word.Letters, letter => Assert.Contains("Times", letter.FontName, StringComparison.OrdinalIgnoreCase));
    }
}

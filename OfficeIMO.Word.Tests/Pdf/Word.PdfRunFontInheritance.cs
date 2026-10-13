using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfRunFontInheritanceTests {
    [Theory]
    [InlineData("body", "document", false)]
    [InlineData("table", "document", false)]
    [InlineData("header", "document", false)]
    [InlineData("footer", "document", false)]
    [InlineData("body", "document", true)]
    [InlineData("table", "document", true)]
    [InlineData("header", "document", true)]
    [InlineData("footer", "document", true)]
    [InlineData("body", "paragraph", false)]
    [InlineData("body", "character", false)]
    [InlineData("table", "table", false)]
    [InlineData("table", "conditional", false)]
    [InlineData("body", "paragraph", true)]
    [InlineData("body", "character", true)]
    [InlineData("table", "table", true)]
    [InlineData("table", "conditional", true)]
    public void UnformattedRunUsesItsOwnInheritanceInsteadOfTheFirstRun(string story, string layer, bool named) {
        using WordDocument document = WordDocument.Create();
        string family = named && layer != "document" ? ManagedTextShapingTestAssets.FamilyName : "Times New Roman";
        document.Settings.FontFamily = layer == "document" ? family : "Courier New";
        WordTable? table = story == "table" ? document.AddTable(1, 1) : null;
        WordParagraph paragraph = story switch {
            "header" => document.HeaderDefaultOrCreate.AddParagraph("LEADING"),
            "footer" => document.FooterDefaultOrCreate.AddParagraph("LEADING"),
            "table" => table!.Rows[0].Cells[0].Paragraphs[0],
            _ => document.AddParagraph("LEADING")
        };
        paragraph.Text = "LEADING";
        paragraph.FontFamily = "Arial";
        WordParagraph inherited = paragraph.AddText(" INHERITED");
        paragraph.AddText(" EXPLICIT").SetFontFamily("Courier New");
        var main = document._wordprocessingDocument.MainDocumentPart!;
        if (layer != "document") {
            var fonts = new RunFonts { Ascii = family, HighAnsi = family };
            var style = new Style(new StyleName { Val = "Inherited run font" }) {
                StyleId = "InheritedRunFont", Type = layer is "table" or "conditional" ? StyleValues.Table :
                    layer == "character" ? StyleValues.Character : StyleValues.Paragraph
            };
            if (layer == "conditional") style.Append(new TableStyleProperties(new RunPropertiesBaseStyle(fonts)) {
                Type = TableStyleOverrideValues.FirstRow
            });
            else style.StyleRunProperties = new StyleRunProperties(fonts);
            main.StyleDefinitionsPart!.Styles!.Append(style);
            if (layer is "table" or "conditional") {
                table!._tableProperties!.TableStyle = new TableStyle { Val = style.StyleId };
                table.ConditionalFormattingFirstRow = true;
            } else if (layer == "character") inherited.SetCharacterStyleId(style.StyleId!);
            else paragraph.SetStyleId(style.StyleId!);
        }
        var options = new PdfOptions();
        if (named) {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
                ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray())));
            if (layer == "document") options.RegisterFontFamilySubstitution(family, ManagedTextShapingTestAssets.FamilyName);
        }
        string original = main.Document.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml +
            string.Concat(main.HeaderParts.Select(part => part.Header.OuterXml)) +
            string.Concat(main.FooterParts.Select(part => part.Footer.OuterXml));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = named ? options : null,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.All(Assert.Single(words, word => word.Text == "INHERITED").Letters, letter =>
            Assert.Contains(named ? "OfficeIMO" : "Times", letter.FontName, StringComparison.OrdinalIgnoreCase));
        Assert.All(Assert.Single(words, word => word.Text == "LEADING").Letters, letter =>
            Assert.Contains("Helvetica", letter.FontName, StringComparison.OrdinalIgnoreCase));
        Assert.All(Assert.Single(words, word => word.Text == "EXPLICIT").Letters, letter =>
            Assert.Contains("Courier", letter.FontName, StringComparison.OrdinalIgnoreCase));
        Assert.Equal(original, main.Document.OuterXml + main.StyleDefinitionsPart.Styles.OuterXml +
            string.Concat(main.HeaderParts.Select(part => part.Header.OuterXml)) +
            string.Concat(main.FooterParts.Select(part => part.Footer.OuterXml)));
    }
}

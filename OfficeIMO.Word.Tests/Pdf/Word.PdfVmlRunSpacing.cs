using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using V = DocumentFormat.OpenXml.Vml;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(50, -10, false)]
    [InlineData(200, 20, false)]
    [InlineData(50, -10, true)]
    [InlineData(200, 20, true)]
    public void VmlCoverTextRetainsDirectAndInheritedWidthAndTracking(int scale, int spacingTwips, bool inherited) {
        using WordDocument document = WordDocument.Create();
        var properties = new RunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = "24" });
        var paragraph = new Paragraph(new Run(properties, new Text("MMMMX")));
        if (inherited) {
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new Style(new StyleRunProperties(new Spacing { Val = spacingTwips }, new CharacterScale { Val = scale })) {
                    StyleId = "VmlRunSpacing", Type = StyleValues.Paragraph
                });
            paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "VmlRunSpacing" });
        } else {
            properties.AddChild(new Spacing { Val = spacingTwips }, true);
            properties.AddChild(new CharacterScale { Val = scale }, true);
        }
        var shape = new V.Shape(new V.TextBox(new TextBoxContent(paragraph))) {
            Id = "VmlSpacingText", Type = "#_x0000_t202",
            Style = "position:absolute;left:72pt;top:72pt;width:360pt;height:120pt"
        };
        document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(new Paragraph(new Run(new Picture(shape)))));
        document.AddParagraph("Body");
        Assert.Empty(document.ValidateDocument());
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value == "M" || letter.Value == "X").ToArray();
        Assert.Equal("MMMMX", string.Concat(letters.Select(letter => letter.Value)));
        double advance = letters[1].StartBaseLine.X - letters[0].StartBaseLine.X;
        Assert.InRange(Math.Abs(advance - (9.996D * scale / 100D + spacingTwips / 20D)), 0D, 0.03D);
    }
}

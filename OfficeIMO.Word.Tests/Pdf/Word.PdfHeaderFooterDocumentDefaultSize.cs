using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false, 12)]
    [InlineData(true, false, 12)]
    [InlineData(false, false, 18)]
    [InlineData(true, false, 18)]
    [InlineData(false, true, 12)]
    [InlineData(true, true, 12)]
    [InlineData(false, true, 18)]
    [InlineData(true, true, 18)]
    public void SaveAsPdf_HeaderFooterUsesEffectiveDocumentDefaultSize(bool footer, bool nativeDoc, int size) {
        using WordDocument source = CreateJoinedParagraphDocument();
        W.Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        string halfPoints = (size * 2).ToString(System.Globalization.CultureInfo.InvariantCulture);
        styles.DocDefaults = new W.DocDefaults(new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
            new W.FontSize { Val = halfPoints }, new W.FontSizeComplexScript { Val = halfPoints })));
        W.Style normal = styles.Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleRunProperties!.RemoveAllChildren<W.FontSize>();
        normal.StyleRunProperties.RemoveAllChildren<W.FontSizeComplexScript>();
        source.AddParagraph("BODY"); source.AddHeadersAndFooters();
        (footer ? (WordHeaderFooter)source.Footer.Default : source.Header.Default).AddParagraph("ALPHA");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(nativeDoc ? WordFileFormat.Doc : WordFileFormat.Docx)));
        using var pdf = OpenJoinedParagraphPdf(document);
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "ALPHA");
        Assert.All(word.Letters, letter => Assert.Equal(size, letter.FontSize, 3));
    }
}

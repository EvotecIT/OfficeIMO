using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordFileFormat.Docx, false)]
    [InlineData(WordFileFormat.Docx, true)]
    [InlineData(WordFileFormat.Doc, false)]
    [InlineData(WordFileFormat.Doc, true)]
    public void SaveAsPdf_AutomaticParagraphShadingClearsInheritedFill(WordFileFormat format, bool useStyle) {
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style normal = styles.Elements<Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleParagraphProperties ??= new StyleParagraphProperties();
        normal.StyleParagraphProperties.Shading = new Shading {
            Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FF0000"
        };
        source.AddParagraph("InheritedRed");
        WordParagraph clear = source.AddParagraph("ClearRed");
        if (useStyle) {
            styles.Append(new Style(new StyleName { Val = "Clear paragraph shading" }, new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" })) {
                StyleId = "ClearParagraphShading", Type = StyleValues.Paragraph, CustomStyle = true
            });
            clear.SetStyleId("ClearParagraphShading");
        } else clear.ShadingFillValue = "auto";

        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(new WordSaveOptions { Format = format })));
        string commands = PdfOperatorSearchText.From(loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes());
        Assert.Single(Regex.Matches(commands, @"(?m)^1 0 0 rg\s*$").Cast<Match>());
    }
}

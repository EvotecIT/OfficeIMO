using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(12)]
    [InlineData(18)]
    public void LegacyDoc_BuiltInHeadingsInheritSourceTypography(int sizePoints) {
        using var document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style normal = styles.Elements<Style>().Single(style => style.StyleId == "Normal");
        normal.StyleRunProperties = new StyleRunProperties(
            new RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial", ComplexScript = "Arial" },
            new FontSize { Val = (sizePoints * 2).ToString() },
            new FontSizeComplexScript { Val = (sizePoints * 2).ToString() },
            new Color { Val = "009977" });
        foreach (WordParagraphStyles heading in new[] { WordParagraphStyles.Heading1, WordParagraphStyles.Heading4, WordParagraphStyles.Heading7, WordParagraphStyles.Heading9 }) {
            document.AddParagraph("InheritedFontMarker" + heading).SetStyle(heading);
            Style style = EnsureParagraphStyle(styles, heading.ToStringStyle());
            style.BasedOn = new BasedOn { Val = "Normal" };
            style.StyleRunProperties = new StyleRunProperties();
        }
        using var reopened = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        Styles importedStyles = reopened._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        foreach (WordParagraphStyles heading in new[] { WordParagraphStyles.Heading1, WordParagraphStyles.Heading4, WordParagraphStyles.Heading7, WordParagraphStyles.Heading9 }) {
            Style style = importedStyles.Elements<Style>().Single(style => style.StyleId == heading.ToStringStyle());
            Assert.Null(style.StyleRunProperties?.GetFirstChild<FontSize>());
            Assert.Null(style.StyleRunProperties?.GetFirstChild<RunFonts>());
            Assert.Null(style.StyleRunProperties?.GetFirstChild<Color>());
        }
        string output = Path.Combine(_directoryWithFiles, "InheritedHeadingTypography" + sizePoints + ".pdf");
        reopened.SaveAsPdf(output, new WordToPdfOptions { IncludePageNumbers = false });
        PdfTextSpan[] spans = PdfReadDocument.Open(File.ReadAllBytes(output)).Pages.SelectMany(page => page.GetTextSpans()).ToArray();
        foreach (WordParagraphStyles heading in new[] { WordParagraphStyles.Heading1, WordParagraphStyles.Heading4, WordParagraphStyles.Heading7, WordParagraphStyles.Heading9 }) {
            PdfTextSpan marker = Assert.Single(spans, span => span.Text.Contains("InheritedFontMarker" + heading));
            Assert.Equal((double)sizePoints, marker.FontSize);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_BuiltInStylesInheritSourceTogglesWithoutTemplateFormatting(bool inheritedEmphasis) {
        string source = Path.Combine(_directoryWithFiles, $"SourceStyleToggles{inheritedEmphasis}.doc");
        using var document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style normal = styles.Elements<Style>().Single(style => style.StyleId == "Normal");
        normal.StyleRunProperties = new StyleRunProperties();
        if (inheritedEmphasis) normal.StyleRunProperties.Append(new Bold(), new Italic());
        foreach (WordParagraphStyles heading in new[] { WordParagraphStyles.Heading1, WordParagraphStyles.Heading4 }) {
            document.AddParagraph("SourceMarker" + heading).SetStyle(heading);
            Style style = EnsureParagraphStyle(styles, heading.ToStringStyle());
            style.BasedOn = new BasedOn { Val = "Normal" };
            style.StyleRunProperties = new StyleRunProperties();
        }
        document.Save(source);
        using var reopened = WordDocument.Load(source);
        string output = source + ".pdf";
        reopened.SaveAsPdf(output, new WordToPdfOptions { IncludePageNumbers = false });
        PdfTextSpan[] spans = PdfReadDocument.Open(File.ReadAllBytes(output)).Pages.SelectMany(page => page.GetTextSpans()).ToArray();
        foreach (WordParagraphStyles heading in new[] { WordParagraphStyles.Heading1, WordParagraphStyles.Heading4 }) {
            PdfTextSpan marker = Assert.Single(spans, span => span.Text.Contains("SourceMarker" + heading));
            Assert.Equal(inheritedEmphasis, marker.IsBold);
            Assert.Equal(inheritedEmphasis, marker.IsItalic);
        }
    }
}

using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void DocumentTheme_ChangedColoursAndFontsReachEveryThemeReference() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            WordStyle accent = document.Styles.AddParagraphStyle("AccentHeading");
            accent.ThemeColor = WordThemeColor.Accent2;
            accent.ThemeFont = WordThemeFont.Major;
            WordParagraph styled = document.AddParagraph("Theme text").SetStyle(accent);
            styled.AddText(" accent run").ThemeColor = WordThemeColor.Accent1;
            styled.Borders.BottomStyle = WordBorderStyle.Single;
            styled.Borders.BottomThemeColor = WordThemeColor.Accent1;

            document.Theme.SetColor(WordThemeColor.Accent1, OfficeColor.Parse("#123456"))
                .SetColor(WordThemeColor.Accent2, OfficeColor.Parse("#ABCDEF"));
            document.Theme.MajorFont = "Georgia";
            document.Theme.MinorFont = "Verdana";
            document.Theme.Name = "Contoso";
            document.Save();
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);
        Assert.Equal("Contoso", reopened.Theme.Name);
        Assert.Equal(OfficeColor.Parse("#123456"), reopened.Theme.GetColor(WordThemeColor.Accent1));
        Assert.Equal("Georgia", reopened.Theme.MajorFont);
        Assert.Equal("Verdana", reopened.Theme.MinorFont);

        WordStyle style = reopened.Styles["AccentHeading"];
        Assert.Equal(OfficeColor.Parse("#ABCDEF"), style.Color);
        Assert.Equal("Georgia", style.FontFamily);
        // The literal values stored beside each theme selection follow the theme for applications that ignore it.
        Assert.Equal("ABCDEF", style.Element.StyleRunProperties!.Color!.Val!.Value);
        Assert.Equal("Georgia", style.Element.StyleRunProperties.RunFonts!.Ascii!.Value);
        Paragraph paragraph = reopened._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single();
        Assert.All(paragraph.Descendants<Color>().Where(color => color.ThemeColor != null), color => Assert.Equal("123456", color.Val!.Value));
        Assert.Equal("123456", paragraph.ParagraphProperties!.ParagraphBorders!.BottomBorder!.Color!.Value);
        Assert.Empty(new OpenXmlValidator().Validate(reopened._wordprocessingDocument));
    }

    [Fact]
    public void DocumentTheme_ImportCopiesTheThemeOfAWordTemplate() {
        using WordDocument template = WordDocument.Load(Path.Combine(_directoryDocuments, "ExampleTemplate.dotx"),
            new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.Styles.AddCharacterStyle("Accent").ThemeColor = WordThemeColor.Accent1;
            document.Theme.Import(template);
            Assert.Throws<InvalidOperationException>(() => template.Theme.MajorFont = "Georgia");
            document.Save();
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);
        Assert.Equal(template.Theme.Name, reopened.Theme.Name);
        Assert.Equal(OfficeColor.Parse("#5B9BD5"), reopened.Theme.GetColor(WordThemeColor.Accent1));
        Assert.Equal(template.Theme.GetColor(WordThemeColor.Text1), reopened.Theme.GetColor(WordThemeColor.Dark1));
        Assert.Equal("Calibri Light", reopened.Theme.MajorFont);
        Assert.Equal("Calibri", reopened.Theme.MinorFont);
        Assert.Equal(OfficeColor.Parse("#5B9BD5"), reopened.Styles["Accent"].Color);
        Assert.Empty(new OpenXmlValidator().Validate(reopened._wordprocessingDocument));
    }
}

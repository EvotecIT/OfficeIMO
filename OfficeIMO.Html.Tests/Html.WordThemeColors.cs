using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void WordToHtml_ThemeColoursAndFontsTakePrecedenceOverTheirCachedValues() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.Theme.SetColor(WordThemeColor.Accent1, OfficeColor.Parse("#123456"));
            document.Theme.MajorFont = "Georgia";
            WordStyle accented = document.Styles.AddParagraphStyle("Accented");
            accented.ThemeColor = WordThemeColor.Accent1;
            accented.ThemeFont = WordThemeFont.Major;
            document.AddParagraph("Styled").SetStyleId("Accented");
            document.AddParagraph("Run").ThemeColor = WordThemeColor.Accent1;
            document.Save();
        }
        // Another producer left stale cached values beside each theme selection.
        using (WordprocessingDocument package = WordprocessingDocument.Open(stream, true)) {
            MainDocumentPart main = package.MainDocumentPart!;
            foreach (W.Color color in main.Document!.Descendants<W.Color>().Concat(main.StyleDefinitionsPart!.Styles!.Descendants<W.Color>())) {
                if (color.ThemeColor != null) color.Val = "FF0000";
            }
            foreach (W.RunFonts fonts in main.StyleDefinitionsPart!.Styles!.Descendants<W.RunFonts>()) {
                if (fonts.AsciiTheme != null) fonts.Ascii = fonts.HighAnsi = "Stale Font";
            }
        }

        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);
        string html = reopened.ToHtml(new WordHtmlExportOptions { IncludeParagraphClasses = true });

        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(html, "color:#123456", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Count);
        Assert.DoesNotContain("ff0000", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Georgia", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Stale Font", html, StringComparison.Ordinal);
    }
}

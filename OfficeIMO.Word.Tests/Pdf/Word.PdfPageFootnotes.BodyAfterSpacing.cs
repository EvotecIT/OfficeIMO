using System.Text.RegularExpressions;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordPdfPageFootnoteTests {
    [Fact]
    public void ImportedAutoSpacingMovesFollowingBodyPastTheReservedFootnotes() {
        const string id = "body-auto-after-spacing";
        string inputPath = PageFootnoteFixturePath(id + ".docx");
        byte[] input = File.ReadAllBytes(inputPath);
        byte[] reference = File.ReadAllBytes(PageFootnoteFixturePath(id + ".microsoft-word.pdf"));
        Assert.Equal("bc4893a2355683c08c310a1e6ccdaa399c45027eb91a7de759010442f44fe72e", PageFootnoteSha256(input));
        Assert.Equal("a2e2722222805655ecd6787a8a328f36f63b90e2d9e06b9335a04e6d66fae729", PageFootnoteSha256(reference));
        using WordDocument source = WordDocument.Load(new MemoryStream(input));
        Assert.Empty(source.ValidateDocument());
        var main = source.OpenXmlDocument.MainDocumentPart!;
        W.Paragraph following = main.Document.Body!.Elements<W.Paragraph>().Last();
        Assert.Null(following.ParagraphProperties);
        W.SpacingBetweenLines inherited = main.StyleDefinitionsPart!.Styles!.DocDefaults!
            .ParagraphPropertiesDefault!.ParagraphPropertiesBaseStyle!.SpacingBetweenLines!;
        Assert.Equal("160", inherited.After!.Value);
        Assert.Equal("259", inherited.Line!.Value);
        Assert.Equal(W.LineSpacingRuleValues.Auto, inherited.LineRule!.Value);
        W.RunFonts defaultFonts = main.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault!
            .RunPropertiesBaseStyle!.GetFirstChild<W.RunFonts>()!;
        Assert.Equal(W.ThemeFontValues.MinorHighAnsi, defaultFonts.AsciiTheme!.Value);
        Assert.Equal(W.ThemeFontValues.MinorHighAnsi, defaultFonts.HighAnsiTheme!.Value);
        string defaultFamily = main.ThemePart!.Theme.ThemeElements!.FontScheme!.MinorFont!.LatinFont!.Typeface!.Value!;
        Assert.Equal("Calibri", defaultFamily);
        string bodyBefore = main.Document.OuterXml;
        string notesBefore = main.FootnotesPart!.Footnotes!.OuterXml;
        string stylesBefore = main.StyleDefinitionsPart.Styles.OuterXml;

        var result = source.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        byte[] actual = result.ToBytes();
        using var pdf = PdfPigDocument.Open(actual);
        using var native = PdfPigDocument.Open(reference);
        string GlyphInventory(PdfPigDocument document) => new(string.Concat(document.GetPages()
                .SelectMany(page => page.Letters).Select(letter => letter.Value))
            .Where(character => !char.IsWhiteSpace(character)).OrderBy(character => character).ToArray());
        Assert.Equal(GlyphInventory(native), GlyphInventory(pdf));
        AssertPageFootnoteTokenOnce(pdf, "FOLLOWING BODY MARKER");
        AssertPageFootnoteTokenOnce(pdf, "NARROW");
        string compact = string.Concat(pdf.GetPages().Select(page => PageFootnoteCompact(page.Text)));
        Assert.True(compact.IndexOf("NARROW", StringComparison.Ordinal) >
            compact.IndexOf("Thequietriver", StringComparison.Ordinal));
        Assert.Equal(2, Regex.Matches(compact, "FINALWORDS\\.").Count);

        var qualification = QualifyPageFootnoteNativeArial(id, actual, result.Warnings,
            declaredSourceFontFamilies: new[] { "Arial", defaultFamily });
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_REQUIRE_NATIVE_ARIAL") == "1")
            Assert.True(qualification.Qualified, "Required native profile unavailable: " + qualification.Qualification);
        if (qualification.Qualified) {
            Assert.Equal(native.NumberOfPages, pdf.NumberOfPages);
            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Equal(2, PageFootnoteTokenPage(pdf, "FOLLOWING BODY MARKER"));
            Assert.Equal(1, PageFootnoteTokenPage(pdf, "NARROW"));
            _output.WriteLine("Native auto-spacing page assignment: qualified and passed.");
        } else {
            _output.WriteLine("Native auto-spacing page assignment: not evaluated; portable inventory/order passed. " +
                qualification.Qualification);
        }
        Assert.Equal(bodyBefore, main.Document.OuterXml);
        Assert.Equal(notesBefore, main.FootnotesPart.Footnotes.OuterXml);
        Assert.Equal(stylesBefore, main.StyleDefinitionsPart.Styles.OuterXml);
        Assert.Equal(input, File.ReadAllBytes(inputPath));
        Assert.Empty(source.ValidateDocument());
    }
}

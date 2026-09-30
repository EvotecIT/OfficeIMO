using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordPdfOmittedFontSizeTests {
    [Theory]
    [InlineData("20", false)]
    [InlineData("24", false)]
    [InlineData("20", true)]
    public void ConditionalLineSpacingUsesTheInheritedTableFontSize(string documentSize, bool derivedStyle) {
        double plainBaseline = FirstTableBaseline(documentSize, includeSpacing: false, derivedStyle);
        double spacedBaseline = FirstTableBaseline(documentSize, includeSpacing: true, derivedStyle);
        // One Arial single-line box at the inherited table size of 11pt.
        Assert.Equal(11D * 1.15D, plainBaseline - spacedBaseline, precision: 3);
    }

    private static double FirstTableBaseline(string documentSize, bool includeSpacing, bool derivedStyle) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.GetFirstChild<FontSize>()!.Val = documentSize;
        var conditional = new TableStyleProperties { Type = TableStyleOverrideValues.FirstRow };
        if (includeSpacing) conditional.Append(new StyleParagraphProperties(
            new SpacingBetweenLines { BeforeLines = 100, Line = "240", LineRule = LineSpacingRuleValues.Auto }));
        styles.Append(new Style(new StyleName { Val = "Sized table" },
            new StyleRunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = derivedStyle ? "24" : "22" }),
            conditional) { Type = StyleValues.Table, StyleId = "SizedTable" });
        if (derivedStyle) styles.Append(new Style(new BasedOn { Val = "SizedTable" },
            new StyleRunProperties(new FontSize { Val = "22" })) { Type = StyleValues.Table, StyleId = "DerivedSizedTable" });
        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = derivedStyle ? "DerivedSizedTable" : "SizedTable" };
        table.ConditionalFormattingFirstRow = true;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Cell";
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var letter = pdf.GetPage(1).Letters.First();
        Assert.Equal(11D, letter.PointSize, precision: 3);
        return letter.StartBaseLine.Y;
    }

    // Reduced templates reproduce Word's fallbacks without importing private documents.
    [Theory]
    [InlineData(false, null, 12D)]
    [InlineData(true, null, 10D)]
    [InlineData(true, "22", 11D)]
    [InlineData(true, "21", 10.5D)]
    public void MissingSizeUsesWordFallbackWhileDeclaredDefaultsRemainExact(bool keepDefaults, string? size, double expectedSize) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        foreach (FontSize fontSize in styles.Descendants<FontSize>().ToArray()) fontSize.Remove();
        foreach (FontSizeComplexScript fontSize in styles.Descendants<FontSizeComplexScript>().ToArray()) fontSize.Remove();
        if (!keepDefaults) styles.DocDefaults?.Remove();
        if (size != null) styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.Append(new FontSize { Val = size });

        document.AddParagraph("Body");
        document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Cell";
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var letters = pdf.GetPages().SelectMany(page => page.Letters).ToArray();
        Assert.Equal("BodyCell", string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => Assert.Equal(expectedSize, letter.PointSize, precision: 3));
    }

    [Fact]
    public void PolicySubstitutionReportsThePolicyRatherThanAMissingFont() {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Arial";
        document.AddParagraph("Policy test");
        using var stream = new System.IO.MemoryStream();
        var report = document.SaveAsPdf(stream, new WordToPdfOptions {
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            IncludePageNumbers = false
        });
        report.RequireSuccess();
        var warning = Assert.Single(report.Warnings, warning =>
            warning.Code == "NativeFontFamilySubstituted" && warning.Details["fontFamily"] == "Arial");
        Assert.Equal("ResourcePolicy", warning.Details["substitutionReason"]);
        Assert.Contains("resource policy", warning.Message, System.StringComparison.Ordinal);
        Assert.DoesNotContain("unavailable", warning.Message, System.StringComparison.Ordinal);
    }
}

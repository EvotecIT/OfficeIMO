using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordPdfOmittedFontSizeTests {
    [Theory]
    [InlineData("20", false, false)]
    [InlineData("24", false, false)]
    [InlineData("20", true, false)]
    [InlineData("20", false, true)]
    public void ConditionalLineSpacingUsesTheInheritedTableFontSize(string documentSize, bool derivedStyle, bool derivedConditional) {
        double plainBaseline = FirstTableBaseline(documentSize, includeSpacing: false, derivedStyle, derivedConditional);
        double spacedBaseline = FirstTableBaseline(documentSize, includeSpacing: true, derivedStyle, derivedConditional);
        // One Arial single-line box at the inherited table size of 11pt.
        Assert.Equal(11D * 1.15D, plainBaseline - spacedBaseline, precision: 3);
    }

    [Theory]
    [InlineData("Times New Roman", "Times", 1.15D)]
    [InlineData("Calibri", "Helvetica", 1.220703125D)]
    public void InheritedConditionalSpacingUsesTheFontFamilyAppliedToTheCell(string family, string pdfFamily, double lineRatio) {
        double plain = FirstTableBaseline("20", false, false, true, family, pdfFamily);
        double spaced = FirstTableBaseline("20", true, false, true, family, pdfFamily);
        Assert.Equal(11D * lineRatio, plain - spaced, precision: 3);
    }

    private static double FirstTableBaseline(string documentSize, bool includeSpacing, bool derivedStyle, bool derivedConditional, string? conditionalFamily = null, string? expectedPdfFamily = null) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.GetFirstChild<FontSize>()!.Val = documentSize;
        var conditional = new TableStyleProperties { Type = TableStyleOverrideValues.FirstRow };
        if (derivedConditional) conditional.Append(new RunPropertiesBaseStyle(new FontSize { Val = "24" }));
        if (includeSpacing) conditional.Append(new StyleParagraphProperties(
            new SpacingBetweenLines { BeforeLines = 100, Line = "240", LineRule = LineSpacingRuleValues.Auto }));
        styles.Append(new Style(new StyleName { Val = "Sized table" },
            new StyleRunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = derivedStyle ? "24" : "22" }),
            conditional) { Type = StyleValues.Table, StyleId = "SizedTable" });
        if (derivedStyle) styles.Append(new Style(new BasedOn { Val = "SizedTable" },
            new StyleRunProperties(new FontSize { Val = "22" })) { Type = StyleValues.Table, StyleId = "DerivedSizedTable" });
        if (derivedConditional) styles.Append(new Style(new BasedOn { Val = "SizedTable" },
            new TableStyleProperties(new RunPropertiesBaseStyle(new FontSize { Val = "22" },
                new RunFonts { Ascii = conditionalFamily, HighAnsi = conditionalFamily })) {
                Type = TableStyleOverrideValues.FirstRow
            }) { Type = StyleValues.Table, StyleId = "DerivedSizedTable" });
        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = derivedStyle || derivedConditional ? "DerivedSizedTable" : "SizedTable" };
        table.ConditionalFormattingFirstRow = true;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Cell";
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var letter = pdf.GetPage(1).Letters.First();
        Assert.Equal(11D, letter.PointSize, precision: 3);
        if (expectedPdfFamily != null) Assert.Contains(expectedPdfFamily, letter.FontName, System.StringComparison.OrdinalIgnoreCase);
        return letter.StartBaseLine.Y;
    }

    // Reduced templates reproduce Word's fallbacks without importing private documents.
    [Theory]
    [InlineData(false, null, 12D, null)]
    [InlineData(true, null, 10D, null)]
    [InlineData(true, "22", 11D, null)]
    [InlineData(true, "21", 10.5D, null)]
    [InlineData(true, null, 10D, "22")]
    [InlineData(true, null, 10D, "21")]
    [InlineData(true, "22", 11D, "26")]
    public void MissingSizeUsesWordFallbackWhileDeclaredDefaultsRemainExact(bool keepDefaults, string? size, double expectedSize, string? complexScriptSize) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        foreach (FontSize fontSize in styles.Descendants<FontSize>().ToArray()) fontSize.Remove();
        foreach (FontSizeComplexScript fontSize in styles.Descendants<FontSizeComplexScript>().ToArray()) fontSize.Remove();
        if (!keepDefaults) styles.DocDefaults?.Remove();
        if (size != null) styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.Append(new FontSize { Val = size });
        if (complexScriptSize != null) styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.Append(new FontSizeComplexScript { Val = complexScriptSize });

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

    [Theory]
    [InlineData("22", null, 11D)]
    [InlineData("21", null, 10.5D)]
    [InlineData("22", "25", 12.5D)]
    public void ComplexScriptRunsUseTheirOwnDeclaredDefaultSize(string defaultSize, string? directSize, double expectedSize) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        foreach (FontSize size in styles.Descendants<FontSize>().ToArray()) size.Remove();
        foreach (FontSizeComplexScript size in styles.Descendants<FontSizeComplexScript>().ToArray()) size.Remove();
        styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.Append(new FontSizeComplexScript { Val = defaultSize });
        WordParagraph body = document.AddParagraph("Body");
        WordParagraph cell = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        cell.Text = "Cell";
        foreach (WordParagraph paragraph in new[] { body, cell }) {
            paragraph._run!.RunProperties ??= new RunProperties();
            paragraph._run.RunProperties.Append(new ComplexScript());
            if (directSize != null) paragraph._run.RunProperties.Append(new FontSizeComplexScript { Val = directSize });
        }
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.Equal(expectedSize, letter.PointSize, precision: 3));
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

    [Theory]
    [InlineData("paragraph", false)]
    [InlineData("character", false)]
    [InlineData("table", false)]
    [InlineData("conditional", false)]
    [InlineData("paragraph", true)]
    [InlineData("character", true)]
    [InlineData("table", true)]
    [InlineData("conditional", true)]
    public void ComplexScriptStylesInheritSizeAndSelectionWhileDirectDisableWins(string layer, bool disable) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        foreach (FontSize size in styles.Descendants<FontSize>().ToArray()) size.Remove();
        foreach (FontSizeComplexScript size in styles.Descendants<FontSizeComplexScript>().ToArray()) size.Remove();
        styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.Append(new FontSizeComplexScript { Val = "22" });
        StyleValues type = layer == "paragraph" ? StyleValues.Paragraph : layer == "character" ? StyleValues.Character : StyleValues.Table;
        var parent = new Style { Type = type, StyleId = "ComplexParent" };
        var child = new Style(new BasedOn { Val = "ComplexParent" }) { Type = type, StyleId = "ComplexChild" };
        if (layer == "conditional") {
            parent.Append(new TableStyleProperties(new RunPropertiesBaseStyle(new FontSizeComplexScript { Val = "28" })) { Type = TableStyleOverrideValues.FirstRow });
            child.Append(new TableStyleProperties(new RunPropertiesBaseStyle(new ComplexScript())) { Type = TableStyleOverrideValues.FirstRow });
        } else {
            parent.Append(new StyleRunProperties(new FontSizeComplexScript { Val = "28" }));
            child.Append(new StyleRunProperties(new ComplexScript()));
        }
        styles.Append(parent, child);
        WordTable table = document.AddTable(1, 1);
        WordParagraph cell = table.Rows[0].Cells[0].Paragraphs[0];
        cell.Text = "Cell";
        cell._run!.RunProperties ??= new RunProperties();
        if (layer == "paragraph") {
            cell._paragraph!.ParagraphProperties ??= new ParagraphProperties();
            cell._paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = "ComplexChild" };
        } else if (layer == "character") {
            cell._run.RunProperties.RunStyle = new RunStyle { Val = "ComplexChild" };
        } else {
            table._tableProperties!.TableStyle = new TableStyle { Val = "ComplexChild" };
            table.ConditionalFormattingFirstRow = true;
        }
        if (disable) cell._run.RunProperties.Append(new ComplexScript { Val = false });
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.Equal(disable ? 10D : 14D, letter.PointSize, precision: 3));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ConditionalCellFontsParticipateInRegistrationAndPolicyDiagnostics(bool enabled, bool trusted) {
        const string family = "OfficeIMO Missing Conditional Font 53EC";
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(new TableStyleProperties(new RunPropertiesBaseStyle(
            new RunFonts { Ascii = family, HighAnsi = family })) { Type = TableStyleOverrideValues.FirstRow }) {
            Type = StyleValues.Table, StyleId = "ConditionalFont"
        });
        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "ConditionalFont" };
        table.ConditionalFormattingFirstRow = enabled;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Cell";
        using var stream = new System.IO.MemoryStream();
        var result = document.SaveAsPdf(stream, new WordToPdfOptions {
            IncludePageNumbers = false,
            ResourcePolicy = trusted ? PdfResourcePolicy.CreateTrustedHost() : PdfResourcePolicy.CreatePortableDeterministic()
        });
        result.RequireSuccess();
        var warnings = result.Warnings.Where(w => w.Code == "NativeFontFamilySubstituted" && w.Details["fontFamily"] == family).ToArray();
        if (enabled) Assert.Equal(trusted ? "FontUnavailable" : "ResourcePolicy", Assert.Single(warnings).Details["substitutionReason"]);
        else Assert.Empty(warnings);
    }
}

using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfDeclaredFontAlternateTests {
    private const string MissingFamily = "OfficeIMO Missing Declared Font";
    private const string FirstAlternate = "OfficeIMO Declared First";
    private const string SecondAlternate = "OfficeIMO Declared Second";

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("table")]
    [InlineData("paragraph-style")]
    [InlineData("character-style")]
    public void DeclaredAlternateSuppliesEffectiveStoryFont(string owner) {
        using WordDocument document = CreateDocument(FirstAlternate);
        AddFontOwner(document, owner);
        string before = SourceXml(document);
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(Options(CreateFonts()));
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "AAAA");
        Assert.Equal(4, word.Letters.Count);
        Assert.Equal(7.2D, word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
        Assert.All(word.Letters, letter => Assert.Contains("OfficeIMO", letter.FontName, StringComparison.Ordinal));
        var warning = Assert.Single(result.Warnings, warning => warning.Code == "NativeFontFamilySubstituted" &&
            warning.Details.TryGetValue("fontFamily", out string? family) && family == MissingFamily);
        Assert.Equal(FirstAlternate, warning.Details["resolvedFontFamily"]);
        Assert.Equal(before, SourceXml(document));
    }

    [Fact]
    public void DeclaredAlternatesKeepFirstAvailableOrderAndIgnoreEmptyDuplicates() {
        using WordDocument document = CreateDocument(",OfficeIMO Missing Alternate,," + FirstAlternate + "," +
            FirstAlternate.ToLowerInvariant() + "," + SecondAlternate + ",");
        AddFontOwner(document, "body");
        var fonts = CreateFonts().RegisterNamedFontFamily(Family(SecondAlternate, 900));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(Options(fonts)));
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "AAAA");
        Assert.Equal(7.2D, word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(false, 3.6D)]
    [InlineData(true, 10.8D)]
    public void ExplicitPrimaryOrCallerMappingPrecedesDeclaredAlternate(bool callerMapping, double expectedAdvance) {
        using WordDocument document = CreateDocument(FirstAlternate);
        AddFontOwner(document, "body");
        var fonts = CreateFonts().RegisterNamedFontFamily(Family(MissingFamily, 300));
        if (callerMapping) {
            fonts.RegisterNamedFontFamily(Family(SecondAlternate, 900));
            fonts.RegisterFontFamilySubstitution(MissingFamily, SecondAlternate);
        }
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(Options(fonts)));
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "AAAA");
        Assert.Equal(expectedAdvance, word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
    }

    [Fact]
    public void DeclaredDefaultAlternateRemainsOnInheritedNamedFontPath() {
        using WordDocument document = CreateDocument(FirstAlternate);
        document.Settings.FontFamily = MissingFamily;
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!
            .RunPropertiesDefault!.RunPropertiesBaseStyle!.RunFonts = Fonts(MissingFamily);
        WordParagraph paragraph = document.AddParagraph("AAAA");
        paragraph.FontSize = 12;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(Options(CreateFonts())));
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "AAAA");
        Assert.Equal(7.2D, word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CallerBodyFontOverridePrecedesDeclaredDocumentDefault(bool converterOverride) {
        using WordDocument document = CreateDocument(FirstAlternate);
        document.Settings.FontFamily = MissingFamily;
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!
            .RunPropertiesDefault!.RunPropertiesBaseStyle!.RunFonts = Fonts(MissingFamily);
        document.AddParagraph("AAAA").FontSize = 12;
        var fonts = new PdfOptions().RegisterNamedFontFamily(Family(FirstAlternate, 300));
        if (!converterOverride) fonts.DefaultFont = PdfStandardFont.Courier;
        WordToPdfOptions options = Options(fonts);
        if (converterOverride) options.FontFamily = "Courier";
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(options));
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "AAAA");
        Assert.All(word.Letters, letter => Assert.Contains("Courier", letter.FontName, StringComparison.Ordinal));
        Assert.Equal(7.2D, word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
    }

    [Fact]
    public void DeclaredLiteralPrimaryIsNotReinterpretedAsACandidateList() {
        const string literalPrimary = "OfficeIMO Declared, Literal";
        using WordDocument document = CreateDocument(FirstAlternate, literalPrimary);
        AddFontOwner(document, "body", literalPrimary);
        var fonts = CreateFonts().RegisterNamedFontFamily(Family(literalPrimary, 300));
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(Options(fonts));
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "AAAA");
        Assert.Equal(3.6D, word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeFontFamilySubstituted" &&
            warning.Details.TryGetValue("fontFamily", out string? family) && family == literalPrimary);
    }

    [Theory]
    [InlineData("OfficeIMO Raw, Literal", "OfficeIMO Raw, Literal", false)]
    [InlineData("OfficeIMO Raw, Literal", "OfficeIMO Raw, Literal", true)]
    [InlineData(@"OfficeIMO Raw\Face", @"OfficeIMO Raw\Face", false)]
    [InlineData(@"OfficeIMO Raw\Face", @"OfficeIMO Raw\Face", true)]
    [InlineData("\"OfficeIMO Raw, Literal\"", "OfficeIMO Raw, Literal", false)]
    [InlineData("\"OfficeIMO Raw, Literal\"", "OfficeIMO Raw, Literal", true)]
    public void DeclaredLiteralAlternatePrecedesParsedFallbackAndPreservesSource(
        string alternateValue, string resolvedFamily, bool documentDefault) {
        using WordDocument document = CreateDocument(alternateValue);
        if (documentDefault) {
            document.Settings.FontFamily = MissingFamily;
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!
                .RunPropertiesDefault!.RunPropertiesBaseStyle!.RunFonts = Fonts(MissingFamily);
            document.AddParagraph("AAAA").FontSize = 12;
        } else AddFontOwner(document, "body");
        var fonts = new PdfOptions().RegisterNamedFontFamily(Family(resolvedFamily, 900))
            .RegisterNamedFontFamily(Family("OfficeIMO Raw", 300));
        string sourceBefore = SourceXml(document);
        string callerBefore = CallerFontState(fonts);
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(Options(fonts));
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "AAAA");
        Assert.Equal(4, word.Letters.Count);
        Assert.Equal(10.8D, word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
        var warning = Assert.Single(result.Warnings, warning => warning.Code == "NativeFontFamilySubstituted" &&
            warning.Details.TryGetValue("fontFamily", out string? family) && family == MissingFamily);
        Assert.Equal(resolvedFamily, warning.Details["resolvedFontFamily"]);
        Assert.Equal(sourceBefore, SourceXml(document));
        Assert.Equal(callerBefore, CallerFontState(fonts));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DeclaredAlternateProfileDistinguishesLiteralFromParsedFallback(bool literalAvailable) {
        const string literalFamily = "OfficeIMO Raw, Literal";
        const string componentFamily = "OfficeIMO Raw";
        using WordDocument document = CreateDocument(literalFamily);
        AddFontOwner(document, "body");
        var faces = new OfficeFontFaceCollection().Add(componentFamily,
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, Enumerable.Range(32, 95).ToArray()));
        if (literalAvailable) faces.Add(literalFamily,
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 900, Enumerable.Range(32, 95).ToArray()));
        var options = Options(new PdfOptions()).UseRenderingProfile(new OfficeRenderingProfile("declared-alternate", faces));
        string sourceBefore = SourceXml(document);
        string callerBefore = CallerFontState(options.PdfOptions!);
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "AAAA");
        Assert.Equal(4, word.Letters.Count);
        Assert.Equal(literalAvailable ? 10.8D : 3.6D,
            word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
        var warning = Assert.Single(result.Warnings, warning => warning.Code == "NativeFontFamilySubstituted" &&
            warning.Details.TryGetValue("fontFamily", out string? family) && family == MissingFamily);
        Assert.Equal(literalAvailable ? literalFamily : componentFamily, warning.Details["resolvedFontFamily"]);
        Assert.Equal(sourceBefore, SourceXml(document));
        Assert.Equal(callerBefore, CallerFontState(options.PdfOptions!));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExactStandardCandidatePrecedesLaterDeclaredFamily(bool standardPrimary) {
        string source = standardPrimary ? "Courier" : MissingFamily;
        using WordDocument document = CreateDocument(standardPrimary ? FirstAlternate : "Courier," + FirstAlternate, source);
        AddFontOwner(document, "body", source);
        var fonts = new PdfOptions().RegisterNamedFontFamily(Family(FirstAlternate, 300));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(Options(fonts)));
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "AAAA");
        Assert.All(word.Letters, letter => Assert.Contains("Courier", letter.FontName, StringComparison.Ordinal));
        Assert.Equal(7.2D, word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData("empty-options", "Courier", 7.2D, true)]
    [InlineData("unrelated-mapping", "Courier", 7.2D, true)]
    [InlineData("explicit-default", "Times-Roman", 8.664D, false)]
    public void StandardDeclaredDefaultDoesNotDependOnUnrelatedCallerConfiguration(
        string configuration, string expectedFont, double expectedAdvance, bool expectSourceWarning) {
        using WordDocument document = CreateDocument("Courier");
        document.Settings.FontFamily = MissingFamily;
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!
            .RunPropertiesDefault!.RunPropertiesBaseStyle!.RunFonts = Fonts(MissingFamily);
        document.AddParagraph("AAAA").FontSize = 12;
        var fonts = new PdfOptions();
        if (configuration != "empty-options") {
            fonts.RegisterNamedFontFamily(Family(FirstAlternate, 600));
            fonts.RegisterFontFamilySubstitution("OfficeIMO Unused Caller Source", FirstAlternate);
        }
        if (configuration == "explicit-default") fonts.DefaultFont = PdfStandardFont.TimesRoman;
        string sourceBefore = SourceXml(document);
        string callerBefore = CallerFontState(fonts);
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(Options(fonts));
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        var word = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "AAAA");
        Assert.Equal(4, word.Letters.Count);
        Assert.All(word.Letters, letter => Assert.Equal(expectedFont, letter.FontName));
        Assert.Equal(expectedAdvance, word.Letters[1].StartBaseLine.X - word.Letters[0].StartBaseLine.X, 3);
        var sourceWarnings = result.Warnings.Where(warning => warning.Code == "NativeFontFamilySubstituted" &&
            warning.Details.TryGetValue("fontFamily", out string? family) && family == MissingFamily).ToArray();
        if (expectSourceWarning) {
            var warning = Assert.Single(sourceWarnings);
            Assert.Equal("Courier", warning.Details["resolvedFontFamily"]);
            Assert.Equal("word:font[" + MissingFamily + "]", warning.Source);
        } else Assert.Empty(sourceWarnings);
        Assert.Equal(sourceBefore, SourceXml(document));
        Assert.Equal(callerBefore, CallerFontState(fonts));
    }

    private static WordDocument CreateDocument(string alternateNames, string primaryFamily = MissingFamily) {
        WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Helvetica";
        document._wordprocessingDocument.MainDocumentPart!.FontTablePart!.Fonts!.Append(
            new W.Font(new W.AltName { Val = alternateNames }) { Name = primaryFamily });
        return document;
    }

    private static void AddFontOwner(WordDocument document, string owner, string sourceFamily = MissingFamily) {
        WordParagraph paragraph;
        if (owner is "header" or "footer") {
            document.AddParagraph("Body");
            document.AddHeadersAndFooters();
            paragraph = (owner == "header" ? (WordHeaderFooter)document.Header.Default : document.Footer.Default)
                .AddParagraph("AAAA");
        } else if (owner == "table") {
            paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "AAAA";
        } else paragraph = document.AddParagraph("AAAA");
        paragraph.FontSize = 12;
        if (owner is "paragraph-style" or "character-style") {
            bool character = owner == "character-style";
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new W.Style(
                new W.StyleName { Val = "Declared Font" }, new W.StyleRunProperties(Fonts(sourceFamily))) {
                Type = character ? W.StyleValues.Character : W.StyleValues.Paragraph, StyleId = "DeclaredFont"
            });
            if (character) (paragraph._run!.RunProperties ??= new W.RunProperties()).RunStyle = new W.RunStyle { Val = "DeclaredFont" };
            else paragraph.SetStyleId("DeclaredFont");
        } else (paragraph._run!.RunProperties ??= new W.RunProperties()).RunFonts = Fonts(sourceFamily);
    }

    private static W.RunFonts Fonts(string family) => new() { Ascii = family, HighAnsi = family };
    private static PdfOptions CreateFonts() => new PdfOptions().RegisterNamedFontFamily(Family(FirstAlternate, 600));
    private static PdfEmbeddedFontFamily Family(string family, ushort advance) => new(family,
        ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, advance, Enumerable.Range(32, 95).ToArray()));
    private static WordToPdfOptions Options(PdfOptions fonts) => new() {
        IncludePageNumbers = false, PdfOptions = fonts, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
    };
    private static string CallerFontState(PdfOptions options) =>
        options.DefaultFont + "|" + options.HeaderFont + "|" + options.FooterFont + "|" +
        string.Join(";", options.NamedFontFamilies.OrderBy(pair => pair.Key, StringComparer.Ordinal).Select(pair =>
            pair.Key + ":" + Convert.ToBase64String(pair.Value.Regular))) + "|" +
        string.Join(";", options.FontFamilySubstitutions.Select(substitution =>
            substitution.SourceFontFamily + ":" + substitution.TargetFontFamily + ":" + substitution.Impact));

    private static string SourceXml(WordDocument document) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        return main.Document.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml + main.FontTablePart!.Fonts!.OuterXml +
            main.DocumentSettingsPart!.Settings!.OuterXml +
            string.Concat(main.HeaderParts.Select(part => part.Header.OuterXml)) +
            string.Concat(main.FooterParts.Select(part => part.Footer.OuterXml));
    }
}

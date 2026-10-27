using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfWord = UglyToad.PdfPig.Content.Word;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordPdfImportedHeadingColorTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ImportedHeadingWithoutDeclaredColorUsesEffectiveDefault(bool header) {
        using WordDocument document = CreateDocument();
        AddImportedHeadingStyle(document, "Heading1");
        WordParagraph paragraph = header
            ? document.HeaderDefaultOrCreate.AddParagraph()
            : document.AddParagraph();
        SetImportedHeading(paragraph, "Heading1", "IMPORTED");
        if (header) {
            AddTopAndBottomPicture(paragraph);
            document.AddParagraph();
        }
        Assert.Empty(document.ValidateDocument());
        string sourceXml = SourceXml(document);

        using var pdf = PdfPigDocument.Open(document.ToPdf(Options()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        AssertRgbWord(page.GetWords(), "IMPORTED", 0, 0, 0);
        Assert.Equal(sourceXml, SourceXml(document));
    }

    [Fact]
    public void SaveAsPdf_ImportedHeadingRetainsDocumentDefaultAndInheritedStyleColors() {
        using WordDocument document = CreateDocument("336699");
        AddImportedHeadingStyle(document, "Heading1");
        W.Styles styles = document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new W.Style(new W.StyleName { Val = "Declared color base" },
            new W.StyleRunProperties(new W.Color { Val = "993366" })) {
            Type = W.StyleValues.Paragraph, StyleId = "DeclaredColorBase"
        });
        AddImportedHeadingStyle(document, "Heading2", basedOn: "DeclaredColorBase");
        SetImportedHeading(document.AddParagraph(), "Heading1", "DEFAULT");
        SetImportedHeading(document.AddParagraph(), "Heading2", "INHERITED");
        Assert.Empty(document.ValidateDocument());
        string sourceXml = SourceXml(document);

        using var pdf = PdfPigDocument.Open(document.ToPdf(Options()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        AssertRgbWord(page.GetWords(), "DEFAULT", 51, 102, 153);
        AssertRgbWord(page.GetWords(), "INHERITED", 153, 51, 102);
        Assert.Equal(sourceXml, SourceXml(document));
    }

    [Fact]
    public void SaveAsPdf_ImportedHeadingRetainsExplicitRunColorAndResetsAutomaticColor() {
        using WordDocument document = CreateDocument();
        AddImportedHeadingStyle(document, "Heading1");
        WordParagraph paragraph = document.AddParagraph();
        paragraph.SetStyleId("Heading1");
        paragraph._paragraph.RemoveAllChildren<W.Run>();
        paragraph._paragraph.Append(
            new W.Run(DirectRunProperties("CC3300"),
                new W.Text("EXPLICIT ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
            new W.Run(DirectRunProperties("auto"),
                new W.Text("AUTO ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
            new W.Run(DirectRunProperties(), new W.Text("DEFAULT")));
        Assert.Empty(document.ValidateDocument());
        string sourceXml = SourceXml(document);

        using var pdf = PdfPigDocument.Open(document.ToPdf(Options()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        AssertRgbWord(page.GetWords(), "EXPLICIT", 204, 51, 0);
        AssertRgbWord(page.GetWords(), "AUTO", 0, 0, 0);
        AssertRgbWord(page.GetWords(), "DEFAULT", 0, 0, 0);
        Assert.Equal(sourceXml, SourceXml(document));
    }

    [Fact]
    public void SaveAsPdf_AuthoredHeadingOneRetainsItsDeclaredDefaultColor() {
        using WordDocument document = CreateDocument();
        WordParagraph heading = document.AddParagraph("AUTHORED").SetStyle(WordParagraphStyles.Heading1);
        heading.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        heading.FontSize = 12;
        W.Style sourceStyle = Assert.Single(document.OpenXmlDocument.MainDocumentPart!
            .StyleDefinitionsPart!.Styles!.Elements<W.Style>(), style => style.StyleId?.Value == "Heading1");
        Assert.Equal("2F5496", sourceStyle.StyleRunProperties!.GetFirstChild<W.Color>()!.Val!.Value);
        Assert.Empty(document.ValidateDocument());
        string sourceXml = SourceXml(document);

        using var pdf = PdfPigDocument.Open(document.ToPdf(Options()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        AssertRgbWord(page.GetWords(), "AUTHORED", 47, 84, 150);
        Assert.Equal(sourceXml, SourceXml(document));
    }

    private static WordDocument CreateDocument(string? defaultColor = null) {
        WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        W.Styles styles = document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        WordSection section = document.Sections[0];
        section.PageSettings.Width = 4800;
        section.PageSettings.Height = 6000;
        section.Margins.Left = section.Margins.Right = 600;
        section.Margins.Top = section.Margins.Bottom = 600;
        section.Margins.HeaderDistance = section.Margins.FooterDistance = 300;
        var runDefaults = new W.RunPropertiesBaseStyle(
            new W.RunFonts { Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName });
        if (defaultColor != null) runDefaults.Append(new W.Color { Val = defaultColor });
        runDefaults.Append(new W.FontSize { Val = "24" });
        styles.DocDefaults = new W.DocDefaults(new W.RunPropertiesDefault(runDefaults));
        W.Style normal = Assert.Single(styles.Elements<W.Style>(), style => style.StyleId?.Value == "Normal");
        normal.StyleRunProperties = new W.StyleRunProperties();
        return document;
    }

    private static void AddImportedHeadingStyle(WordDocument document, string id, string? basedOn = null) {
        W.Styles styles = document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        foreach (W.Style existing in styles.Elements<W.Style>().Where(style => style.StyleId?.Value == id).ToArray())
            existing.Remove();
        var style = new W.Style(new W.StyleName { Val = id }) {
            Type = W.StyleValues.Paragraph, StyleId = id, CustomStyle = true
        };
        if (basedOn != null) style.Append(new W.BasedOn { Val = basedOn });
        style.Append(new W.StyleParagraphProperties(
            new W.SpacingBetweenLines { Before = "0", After = "0", Line = "300", LineRule = W.LineSpacingRuleValues.Exact },
            new W.OutlineLevel { Val = id == "Heading2" ? 1 : 0 }));
        document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
    }

    private static void SetImportedHeading(WordParagraph paragraph, string id, string text) {
        paragraph.SetStyleId(id);
        paragraph._paragraph.RemoveAllChildren<W.Run>();
        paragraph._paragraph.Append(new W.Run(DirectRunProperties(), new W.Text(text)));
    }

    private static W.RunProperties DirectRunProperties(string? color = null) {
        var properties = new W.RunProperties(new W.RunFonts {
            Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName
        });
        if (color != null) properties.Append(new W.Color { Val = color });
        properties.Append(new W.FontSize { Val = "24" });
        return properties;
    }

    private static void AddTopAndBottomPicture(WordParagraph paragraph) {
        // Match the admitted native heading-color control's source profile:
        // 120x32pt, column/paragraph offsets0, four explicit8pt distances.
        using var png = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(4, 2));
        WordImage image = paragraph.AddText(string.Empty).AddImage(png, "heading.png", OfficeIMO.OfficeLength.Pixels(160D), OfficeIMO.OfficeLength.Pixels(128D / 3D), options: new OfficeIMO.Word.WordImageOptions { Wrapping = WordImageTextWrapping.TopAndBottom });
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Column;
        image.HorizontalPositionOffset = OfficeIMO.OfficeLength.Emus(0L);
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Paragraph;
        image.VerticalPositionOffset = OfficeIMO.OfficeLength.Emus(0L);
        DW.Anchor anchor = image._Image.Anchor!;
        anchor.DistanceFromLeft = anchor.DistanceFromTop = anchor.DistanceFromRight = anchor.DistanceFromBottom = 101600U;
        anchor.BehindDoc = false;
        foreach (DW.EffectExtent extent in anchor.Descendants<DW.EffectExtent>())
            extent.LeftEdge = extent.TopEdge = extent.RightEdge = extent.BottomEdge = 0L;
    }

    private static WordPdfExportOptions Options() {
        var pdf = new PdfCreateOptions();
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500,
            Enumerable.Range(32, 95).ToArray());
        pdf.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
            font, bold: font, italic: font, boldItalic: font));
        return new WordPdfExportOptions {
            IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = pdf
        };
    }

    private static void AssertRgbWord(IEnumerable<PdfWord> words, string text, byte red, byte green, byte blue) {
        PdfWord word = Assert.Single(words, candidate => candidate.Text == text);
        Assert.Equal(text, string.Concat(word.Letters.Select(letter => letter.Value)));
        Assert.All(word.Letters, letter => {
            Assert.Contains("OfficeIMOShapingTest", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(12D, letter.FontSize, 3);
            var (actualRed, actualGreen, actualBlue) = letter.Color.ToRGBValues();
            Assert.Equal((int)red, (int)Math.Round(actualRed * 255D));
            Assert.Equal((int)green, (int)Math.Round(actualGreen * 255D));
            Assert.Equal((int)blue, (int)Math.Round(actualBlue * 255D));
        });
    }

    private static string SourceXml(WordDocument document) {
        var main = document.OpenXmlDocument.MainDocumentPart!;
        return main.Document.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml +
            main.DocumentSettingsPart!.Settings!.OuterXml +
            string.Concat(main.HeaderParts.Select(part => part.Header.OuterXml)) +
            string.Concat(main.FooterParts.Select(part => part.Footer.OuterXml));
    }
}

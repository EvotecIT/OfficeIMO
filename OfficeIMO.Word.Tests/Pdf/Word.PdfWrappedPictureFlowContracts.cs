using System.Text.Json;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordPdfWrappedPictureFlowContractTests {
    private const double Line = 16D, ImageWidth = 120D, ImageHeight = 32D, Distance = Line / 2D;

    [Theory]
    [InlineData(WordImageTextWrapping.TopAndBottom)]
    [InlineData(WordImageTextWrapping.Square)]
    public void WrappedKeepWithNextGroupMovesTogetherWhenOnlyTheWrappedParagraphFits(WordImageTextWrapping wrapping) {
        double paragraphHeight = wrapping == WordImageTextWrapping.TopAndBottom ? ImageHeight + Distance + Line : 2D * Line;
        double remaining = paragraphHeight + Line / 2D;
        Assert.True(2D * Line <= remaining && paragraphHeight <= remaining && paragraphHeight + Line > remaining);
        using var control = KeepDocument(WordImageTextWrapping.InFrontOfText, remaining);
        var controlResult = control.ToPdfDocumentResult(Options());
        byte[] controlBytes = Save("keep-control-" + wrapping, control, controlResult);
        using var controlPdf = PdfPigDocument.Open(controlBytes);
        Assert.Equal(1, TokenPage(controlPdf, "ANCHOR"));
        Assert.Equal(TokenPage(controlPdf, "ANCHOR"), TokenPage(controlPdf, "TAIL"));
        using var word = KeepDocument(wrapping, remaining);
        string source = SourceXml(word);
        Assert.Empty(word.ValidateDocument());
        var result = word.ToPdfDocumentResult(Options());
        byte[] bytes = Save("keep-" + wrapping, word, result);
        Assert.Equal(source, SourceXml(word));
        AssertAdmitted(result);
        using var pdf = PdfPigDocument.Open(bytes);
        int anchorPage = TokenPage(pdf, "ANCHOR"), tailPage = TokenPage(pdf, "TAIL");
        Assert.Equal(anchorPage, tailPage);
        Assert.True(anchorPage > 1, "The complete keep group must move to an unobstructed later page.");
        Assert.Single(pdf.GetPage(anchorPage).GetImages());
    }

    [Theory]
    [InlineData(false, WordImageTextWrapping.TopAndBottom)]
    [InlineData(false, WordImageTextWrapping.Square)]
    [InlineData(true, WordImageTextWrapping.TopAndBottom)]
    public void HeadingPictureUsesTheSameExclusionsAsParagraphText(bool header, WordImageTextWrapping wrapping) {
        using var word = BaseDocument();
        WordParagraph heading;
        if (header) {
            word.AddHeadersAndFooters();
            var story = word.Sections[0].Header.Default!;
            heading = story.Paragraphs.FirstOrDefault() ?? story.AddParagraph();
        } else heading = word.AddParagraph();
        heading.Style = WordParagraphStyles.Heading1;
        Configure(heading, "HEADING");
        heading.KeepWithNextOverride = false;
        AddImage(word, heading, wrapping);
        Configure(word.AddParagraph(), "BODY");
        string source = SourceXml(word);
        Assert.Empty(word.ValidateDocument());
        var result = word.ToPdfDocumentResult(Options());
        byte[] bytes = Save("heading-" + header + "-" + wrapping, word, result);
        Assert.Equal(source, SourceXml(word));
        AssertAdmitted(result);
        using var pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(1);
        var image = Assert.Single(page.GetImages()).BoundingBox;
        var text = Assert.Single(page.GetWords(), item => item.Text == "HEADING").BoundingBox;
        if (wrapping == WordImageTextWrapping.TopAndBottom)
            Assert.True(text.Top <= image.Bottom - Distance + .01D, "TopAndBottom heading text must clear the source picture and its bottom distance.");
        else {
            bool alongside = text.Top > image.Bottom - Distance && text.Bottom < image.Top + Distance;
            Assert.True(!alongside || text.Right <= image.Left - Distance + .01D || text.Left >= image.Right + Distance - .01D,
                "A Square heading must not paint inside its picture exclusion.");
        }
        if (header) {
            var body = Assert.Single(page.GetWords(), item => item.Text == "BODY").BoundingBox;
            Assert.True(body.Top <= text.Bottom + .01D,
                "Body reservation must use the expanded heading rows that actually paint.");
        }
    }

    [Fact]
    public void WrappedKeepTogetherMovesItsPictureWithTheAdmittedParagraph() {
        double wrappedHeight = ImageHeight + Distance + Line;
        using var word = KeepDocument(WordImageTextWrapping.TopAndBottom, wrappedHeight - Line / 2D);
        WordParagraph anchor = word.Paragraphs.Single(paragraph => paragraph.Text == "ANCHOR WORD TWO");
        anchor.KeepWithNextOverride = false;
        Assert.Empty(word.ValidateDocument());
        string source = SourceXml(word);
        var result = word.ToPdfDocumentResult(Options());
        using var pdf = PdfPigDocument.Open(Save("keep-together", word, result));
        Assert.Equal(source, SourceXml(word));
        AssertAdmitted(result);
        int page = TokenPage(pdf, "ANCHOR");
        Assert.True(page > 1);
        Assert.Equal(page, TokenPage(pdf, "TWO"));
        Assert.Single(pdf.GetPage(page).GetImages());
        Assert.Empty(pdf.GetPage(1).GetImages());
    }

    [Fact]
    public void WrappedWidowControlMovesThreePhysicalRowsAndTheirPictureTogether() {
        double remaining = ImageHeight + Distance + 2.5D * Line;
        using var word = KeepDocument(WordImageTextWrapping.TopAndBottom, remaining);
        WordParagraph anchor = word.Paragraphs.Single(paragraph => paragraph.Text == "ANCHOR WORD TWO");
        anchor.Text = "ANCHOR\nMIDDLE\nLAST";
        anchor.KeepWithNextOverride = false;
        anchor.KeepLinesTogether = false;
        anchor.AvoidWidowAndOrphan = true;
        Assert.Empty(word.ValidateDocument());
        string source = SourceXml(word);
        var result = word.ToPdfDocumentResult(Options());
        using var pdf = PdfPigDocument.Open(Save("widow-rows", word, result));
        Assert.Equal(source, SourceXml(word));
        AssertAdmitted(result);
        int page = TokenPage(pdf, "ANCHOR");
        Assert.True(page > 1);
        Assert.Equal(page, TokenPage(pdf, "MIDDLE"));
        Assert.Equal(page, TokenPage(pdf, "LAST"));
        Assert.Single(pdf.GetPage(page).GetImages());
        Assert.Empty(pdf.GetPage(1).GetImages());
    }

    private static WordDocument KeepDocument(WordImageTextWrapping wrapping, double remaining) {
        var word = BaseDocument();
        double fillerHeight = 240D - remaining;
        int lines = (int)Math.Floor(fillerHeight / Line);
        WordParagraph? last = null;
        for (int index = 0; index < lines; index++) Configure(last = word.AddParagraph(), "FILL" + index);
        if (last != null) last.LineSpacingAfterPoints = fillerHeight - lines * Line;
        var anchor = word.AddParagraph();
        Configure(anchor, "ANCHOR WORD TWO");
        anchor.KeepWithNext = true;
        anchor.KeepLinesTogether = true;
        AddImage(word, anchor, wrapping);
        Configure(word.AddParagraph(), "TAIL");
        return word;
    }

    private static WordDocument BaseDocument() {
        var word = WordDocument.Create();
        word.Settings.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        var section = word.Sections[0];
        section.PageSettings.Width = 4800; section.PageSettings.Height = 6000;
        section.Margins.Left = section.Margins.Right = 600;
        section.Margins.Top = section.Margins.Bottom = 600;
        section.Margins.HeaderDistance = section.Margins.FooterDistance = 360;
        return word;
    }

    private static void Configure(WordParagraph paragraph, string text) {
        paragraph.Text = text;
        paragraph.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        paragraph.FontSize = 12;
        paragraph.LineSpacingPoints = Line;
        paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
        paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0D;
        paragraph.AvoidWidowAndOrphan = false;
    }

    private static void AddImage(WordDocument word, WordParagraph paragraph, WordImageTextWrapping wrapping) {
        byte[] png = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(2, 2, OfficeColor.Black), OfficeImageExportFormat.Png);
        using var stream = new MemoryStream(png);
        var image = paragraph.AddText(string.Empty).InsertImage(stream, "picture.png", ImageWidth / .75D, ImageHeight / .75D, wrapping);
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        image.HorizontalPositionOffset = 30L * 12700L;
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Paragraph;
        image.VerticalPositionOffset = 0L;
        var main = word.OpenXmlDocument!.MainDocumentPart!;
        var anchor = main.Document!.Descendants<DW.Anchor>()
            .Concat(main.HeaderParts.SelectMany(part => part.Header!.Descendants<DW.Anchor>())).Single();
        anchor.DistanceFromLeft = anchor.DistanceFromRight = anchor.DistanceFromTop = anchor.DistanceFromBottom = (uint)(Distance * 12700D);
        foreach (var extent in anchor.Descendants<DW.EffectExtent>()) extent.LeftEdge = extent.TopEdge = extent.RightEdge = extent.BottomEdge = 0L;
    }

    private static WordToPdfOptions Options() {
        var pdf = new PdfOptions();
        pdf.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, Enumerable.Range(32, 95).ToArray())));
        return new WordToPdfOptions { IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = pdf };
    }

    private static int TokenPage(PdfPigDocument pdf, string token) => Assert.Single(pdf.GetPages(), page => page.GetWords().Any(word => word.Text == token)).Number;
    private static void AssertAdmitted(PdfDocumentConversionResult result) => Assert.DoesNotContain(result.Warnings,
        warning => warning.Code is "NativeAnchoredImageFlowed" or "NativeAnchoredImageWrapApproximated");
    private static string SourceXml(WordDocument word) {
        var main = word.OpenXmlDocument!.MainDocumentPart!;
        return main.Document!.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml +
            string.Concat(main.HeaderParts.Select(part => part.Header!.OuterXml));
    }
    private static byte[] Save(string name, WordDocument word, PdfDocumentConversionResult result) {
        byte[] bytes = result.ToBytes();
        string output = Path.Combine(AppContext.BaseDirectory, "picture-wrap-flow-artifacts"); Directory.CreateDirectory(output);
        File.WriteAllBytes(Path.Combine(output, name + ".pdf"), bytes);
        File.WriteAllText(Path.Combine(output, name + ".source.xml"), SourceXml(word));
        using var pdf = PdfPigDocument.Open(bytes);
        var info = new { Pages = pdf.GetPages().Select(page => new { page.Number,
            Words = page.GetWords().Select(item => new { item.Text, item.BoundingBox.Left, item.BoundingBox.Right, item.BoundingBox.Top, item.BoundingBox.Bottom }),
            Images = page.GetImages().Select(item => new { item.BoundingBox.Left, item.BoundingBox.Right, item.BoundingBox.Top, item.BoundingBox.Bottom }) }),
            Warnings = result.Warnings.Select(item => new { item.Code, item.Message, item.LossKind }) };
        File.WriteAllText(Path.Combine(output, name + ".geometry.json"), JsonSerializer.Serialize(info, new JsonSerializerOptions { WriteIndented = true }));
        return bytes;
    }
}

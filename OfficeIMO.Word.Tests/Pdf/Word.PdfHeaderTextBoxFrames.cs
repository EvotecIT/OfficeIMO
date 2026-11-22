using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfWord = UglyToad.PdfPig.Content.Word;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    // An anchored header text box keeps its own page position and width, as in Word: its text
    // starts at the box's left inset and wraps inside the box, while the header line stays put.
    [Fact]
    public void SaveAsPdf_HeaderTextBoxKeepsItsPagePositionAndWidth() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordHeader header = document.PageSetup.Header;
        header.AddParagraph("HeaderLine");
        WordTextBox box = header.AddTextBox("Boxed words wrap inside the header text box across lines", WordImageTextWrapping.InFrontOfText);
        box.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        box.HorizontalPositionOffset = global::OfficeIMO.OfficeLength.Points(300);
        box.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        box.VerticalPositionOffset = global::OfficeIMO.OfficeLength.Points(24);
        box.Width = global::OfficeIMO.OfficeLength.Points(150);
        box.Height = global::OfficeIMO.OfficeLength.Points(80);
        box.RelativeWidthPercentage = 0;
        box.RelativeHeightPercentage = 0;

        using var pdf = PdfPigDocument.Open(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Page page = pdf.GetPage(1);
        PdfWord[] boxed = page.GetWords().Where(word => word.Text is "Boxed" or "lines").ToArray();
        Assert.Equal(2, boxed.Length);
        // The box text starts inside the box and wraps onto a lower line within its 150pt width.
        Assert.InRange(boxed[0].Letters[0].StartBaseLine.X, 300D, 310D);
        Assert.True(boxed[1].Letters[0].StartBaseLine.Y < boxed[0].Letters[0].StartBaseLine.Y - 1D);
        Assert.All(page.Letters.Where(letter => letter.StartBaseLine.Y > 792D - 104D && letter.StartBaseLine.X >= 300D),
            letter => Assert.InRange(letter.EndBaseLine.X, 300D, 450.5D));
        double top = 792D - 24D;
        Assert.InRange(boxed[0].Letters[0].StartBaseLine.Y, top - 80D, top);
        Assert.InRange(page.GetWords().Single(word => word.Text == "HeaderLine").Letters[0].StartBaseLine.X, 71.5D, 72.5D);
    }

    // A hyperlink inside a header box is clickable where its text is drawn, inside the box.
    [Fact]
    public void SaveAsPdf_HeaderTextBoxHyperlinkCoversItsText() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordTextBox box = document.PageSetup.Header.AddTextBox("Visit ", WordImageTextWrapping.InFrontOfText);
        box.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        box.HorizontalPositionOffset = global::OfficeIMO.OfficeLength.Points(300);
        box.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        box.VerticalPositionOffset = global::OfficeIMO.OfficeLength.Points(24);
        box.Width = global::OfficeIMO.OfficeLength.Points(150);
        box.Height = global::OfficeIMO.OfficeLength.Points(40);
        box.RelativeWidthPercentage = 0;
        box.RelativeHeightPercentage = 0;
        box.Paragraphs.Single().AddHyperLink("LINKTEXT", new Uri("https://example.com/"));

        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Page page = pdf.GetPage(1);
        var link = Assert.Single(page.GetAnnotations(), annotation => annotation.Type == UglyToad.PdfPig.Annotations.AnnotationType.Link);
        Letter first = page.Letters.First(letter => letter.Value == "L");
        Assert.InRange(first.StartBaseLine.X, link.Rectangle.Left - 0.5D, link.Rectangle.Right);
        Assert.InRange(first.StartBaseLine.Y, link.Rectangle.Bottom - 0.5D, link.Rectangle.Top);
        Assert.InRange(link.Rectangle.Top, 792D - 64D, 792D - 24D);
    }

    // Header text wraps around a square-wrapped box beside it rather than running under it.
    [Fact]
    public void SaveAsPdf_HeaderTextWrapsAroundASquareWrappedTextBox() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordHeader header = document.PageSetup.Header;
        WordParagraph line = header.AddParagraph("Header words keep clear of the box and wrap beside it on the right side of the page");
        WordTextBox box = line.AddTextBox("Box", WordImageTextWrapping.Square);
        box.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        box.HorizontalPositionOffset = global::OfficeIMO.OfficeLength.Points(72);
        box.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        box.VerticalPositionOffset = global::OfficeIMO.OfficeLength.Points(30);
        box.Width = global::OfficeIMO.OfficeLength.Points(150);
        box.Height = global::OfficeIMO.OfficeLength.Points(60);
        box.RelativeWidthPercentage = 0;
        box.RelativeHeightPercentage = 0;

        using var pdf = PdfPigDocument.Open(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Page page = pdf.GetPage(1);
        PdfWord first = page.GetWords().First(word => word.Text == "Header");
        Assert.True(first.Letters[0].StartBaseLine.X >= 222D, $"Header text starts at {first.Letters[0].StartBaseLine.X}.");
    }

    // Word paints the header layer before the body, so a filled box in front of the header text that
    // reaches into the body area still sits under the body text.
    [Fact]
    public void SaveAsPdf_HeaderTextBoxOverBodyPaintsUnderTheBodyText() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("BODYTEXT runs under the header box");
        WordTextBox box = document.PageSetup.Header.AddTextBox("BOXTEXT", WordImageTextWrapping.InFrontOfText);
        box.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        box.HorizontalPositionOffset = global::OfficeIMO.OfficeLength.Points(60);
        box.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        box.VerticalPositionOffset = global::OfficeIMO.OfficeLength.Points(60);
        box.Width = global::OfficeIMO.OfficeLength.Points(300);
        box.Height = global::OfficeIMO.OfficeLength.Points(40);
        box.RelativeWidthPercentage = 0;
        box.RelativeHeightPercentage = 0;
        box.FillColor = global::OfficeIMO.OfficeColor.Parse("FFFF00");

        var result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Page page = pdf.GetPage(1);
        Assert.Contains(page.Paths, path => path.IsFilled && path.FillColor.ToRGBValues() == (1D, 1D, 0D) &&
            path.GetBoundingRectangle() is { } bounds && Math.Abs(bounds.Left - 60D) < 0.01D && Math.Abs(bounds.Top - (792D - 60D)) < 0.01D);
        Letter[] letters = page.Letters.ToArray();
        int boxText = Array.FindIndex(letters, letter => letter.Value == "X");
        int bodyText = Array.FindIndex(letters, letter => letter.Value == "D");
        Assert.InRange(letters[boxText - 2].StartBaseLine.X, 67D, 68.5D);
        Assert.True(boxText < bodyText, "The body text paints after the header box.");
    }

    // Box paragraphs use the ordinary paragraph route, so their shading paints inside the frame.
    [Fact]
    public void SaveAsPdf_HeaderTextBoxParagraphShadingPaintsInsideTheBox() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordTextBox box = document.PageSetup.Header.AddTextBox("Shaded box text", WordImageTextWrapping.InFrontOfText);
        box.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        box.HorizontalPositionOffset = global::OfficeIMO.OfficeLength.Points(300);
        box.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        box.VerticalPositionOffset = global::OfficeIMO.OfficeLength.Points(20);
        box.Width = global::OfficeIMO.OfficeLength.Points(150);
        box.Height = global::OfficeIMO.OfficeLength.Points(40);
        box.RelativeWidthPercentage = 0;
        box.RelativeHeightPercentage = 0;
        box.Paragraphs.Single().ShadingFillColor = global::OfficeIMO.OfficeColor.Parse("FF0000");

        var result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains(pdf.GetPage(1).Paths, path => path.IsFilled && path.FillColor.ToRGBValues() == (1D, 0D, 0D) &&
            path.GetBoundingRectangle() is { } bounds && bounds.Left >= 300D && bounds.Right <= 450.01D && bounds.Top <= 792D - 20D);
    }

    // A word space that ends a line after superscript text is drawn in its own run's size, as Word does.
    [Fact]
    public void SaveAsPdf_LineEndingSpaceAfterSuperscriptKeepsItsOwnSize() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph(new string('a', 89) + "x");
        paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(10);
        paragraph.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        foreach (string text in new[] { "2", " next words" }) {
            WordParagraph run = paragraph.AddText(text);
            run.FontSize = global::OfficeIMO.OfficeLength.Points(10);
            run.FontFamily = ManagedTextShapingTestAssets.FamilyName;
            if (text == "2") run.SetSuperscript();
        }

        using var pdf = PdfPigDocument.Open(RenderWithFont(document,
            ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(800, -500, 0, 500, Enumerable.Range(32, 95).ToArray())));
        Letter[] letters = pdf.GetPage(1).Letters.ToArray();
        int script = Array.FindIndex(letters, letter => letter.Value == "2");
        Assert.True(script >= 0 && script + 1 < letters.Length);
        Assert.Equal(" ", letters[script + 1].Value);
        Assert.True(letters[script].PointSize < 9D);
        Assert.Equal(10D, letters[script + 1].PointSize, 2);
        Assert.True(letters.Single(letter => letter.Value == "n").StartBaseLine.Y < letters[script].StartBaseLine.Y - 5D);
    }

    private static WordTextBox AddPlacedHeaderBox(WordParagraph host, string text, WordImageTextWrapping wrap,
        double left = 300D, double top = 24D, double width = 150D, double height = 60D) {
        WordTextBox box = host.AddTextBox(text, wrap);
        box.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        box.HorizontalPositionOffset = global::OfficeIMO.OfficeLength.Points(left);
        box.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        box.VerticalPositionOffset = global::OfficeIMO.OfficeLength.Points(top);
        box.Width = global::OfficeIMO.OfficeLength.Points(width);
        box.Height = global::OfficeIMO.OfficeLength.Points(height);
        box.RelativeWidthPercentage = 0;
        box.RelativeHeightPercentage = 0;
        return box;
    }

    // A frame holds one page of paragraphs, so a page break inside the box keeps the flowed layout.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_HeaderTextBoxWithPageBreakBeforeKeepsFlowedLayout(bool inherited) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordTextBox box = AddPlacedHeaderBox(document.PageSetup.Header.AddParagraph("Header"), "First", WordImageTextWrapping.InFrontOfText);
        W.Paragraph second = (W.Paragraph)box.Content!.Elements<W.Paragraph>().Single().CloneNode(true);
        second.Descendants<W.Text>().Single().Text = "Second";
        second.Descendants<W.Break>().ToList().ForEach(lineBreak => lineBreak.Remove());
        second.ParagraphProperties ??= new W.ParagraphProperties();
        if (inherited) {
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new W.Style(
                new W.StyleParagraphProperties(new W.PageBreakBefore())) { Type = W.StyleValues.Paragraph, StyleId = "BreakBefore" });
            second.ParagraphProperties.ParagraphStyleId = new W.ParagraphStyleId { Val = "BreakBefore" };
        } else second.ParagraphProperties.PageBreakBefore = new W.PageBreakBefore();
        box.Content.Append(second);

        var result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        string text = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("First", text, StringComparison.Ordinal);
        Assert.Contains("Second", text, StringComparison.Ordinal);
    }

    // Inner indents that leave no text width keep the flowed layout instead of failing the export.
    [Fact]
    public void SaveAsPdf_HeaderTextBoxWithOversizedIndentsKeepsFlowedLayout() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordTextBox box = AddPlacedHeaderBox(document.PageSetup.Header.AddParagraph("Header"), "Indented", WordImageTextWrapping.InFrontOfText, width: 75D);
        WordParagraph inner = box.Paragraphs.Single();
        inner.IndentationBefore = global::OfficeIMO.OfficeLength.Points(45);
        inner.IndentationAfter = global::OfficeIMO.OfficeLength.Points(25);

        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains("Indented", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)), StringComparison.Ordinal);
    }

    // The outline narrows the box's text area too, so indents that only fit without it keep the flowed layout.
    [Fact]
    public void SaveAsPdf_OutlinedHeaderTextBoxCountsTheOutlineAgainstItsIndents() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordTextBox box = AddPlacedHeaderBox(document.PageSetup.Header.AddParagraph("Header"), "Indented", WordImageTextWrapping.InFrontOfText, width: 75D);
        var properties = box.DrawingShapeProperties!;
        properties.RemoveAllChildren<A.Outline>();
        properties.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })) { Width = 25400 });
        // 75pt less the 7.2pt default insets leaves 0.6pt beside 45pt and 15pt indents; the 2pt outline takes the rest.
        WordParagraph inner = box.Paragraphs.Single();
        inner.IndentationBefore = global::OfficeIMO.OfficeLength.Points(45);
        inner.IndentationAfter = global::OfficeIMO.OfficeLength.Points(15);

        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains("Indented", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)), StringComparison.Ordinal);
    }

    // As in Word, a box in front of text paints over the header text it overlaps, including later header
    // paragraphs, while a box behind text stays under it.
    [Theory]
    [InlineData(WordImageTextWrapping.InFrontOfText)]
    [InlineData(WordImageTextWrapping.BehindText)]
    public void SaveAsPdf_HeaderTextBoxStacksAgainstTheHeaderText(WordImageTextWrapping wrap) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordHeader header = document.PageSetup.Header;
        WordTextBox box = AddPlacedHeaderBox(header.AddParagraph("ANCHORLINE"), "QQQ", wrap, left: 60D, top: 30D, width: 200D, height: 60D);
        box.FillColor = global::OfficeIMO.OfficeColor.Parse("FFFF00");
        header.AddParagraph("FOLLOWINGLINE");

        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        string order = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));
        int boxText = order.IndexOf("QQQ", StringComparison.Ordinal);
        int anchorText = order.IndexOf("ANCHORLINE", StringComparison.Ordinal);
        int followingText = order.IndexOf("FOLLOWINGLINE", StringComparison.Ordinal);
        int bodyText = order.IndexOf("Body", StringComparison.Ordinal);
        Assert.True(boxText >= 0 && anchorText >= 0 && followingText >= 0 && bodyText >= 0, order);
        if (wrap == WordImageTextWrapping.InFrontOfText)
            Assert.True(boxText > anchorText && boxText > followingText, $"The box paints after the header text: {order}");
        else
            Assert.True(boxText < anchorText && boxText < followingText, $"The box paints under the header text: {order}");
        Assert.True(boxText < bodyText, $"The header layer paints before the body: {order}");
    }

    // A square-wrapped box keeps header text clear of its rectangle even when it paints nothing.
    [Fact]
    public void SaveAsPdf_UnpaintedSquareWrappedHeaderTextBoxStillPushesTextAside() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordParagraph line = document.PageSetup.Header.AddParagraph("Header words keep clear of the box and wrap beside it on the right side of the page");
        WordTextBox box = AddPlacedHeaderBox(line, "Box", WordImageTextWrapping.Square, left: 72D, top: 30D, height: 60D);
        var properties = box.DrawingShapeProperties!;
        properties.RemoveAllChildren<A.SolidFill>();
        properties.RemoveAllChildren<A.Outline>();
        properties.Append(new A.NoFill(), new A.Outline(new A.NoFill()));
        W.Paragraph inner = box.Content!.Elements<W.Paragraph>().Single();
        inner.RemoveAllChildren<W.Run>();
        inner.ParagraphProperties ??= new W.ParagraphProperties();
        inner.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties(new W.Vanish());

        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        PdfWord first = pdf.GetPage(1).GetWords().First(word => word.Text == "Header");
        Assert.True(first.Letters[0].StartBaseLine.X >= 222D, $"Header text starts at {first.Letters[0].StartBaseLine.X}.");
    }

    // Links in content past the bottom of a box are not clickable; a link crossing the edge is trimmed to it.
    [Fact]
    public void SaveAsPdf_HeaderTextBoxLinksStayInsideTheVisibleFrame() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordTextBox box = AddPlacedHeaderBox(document.PageSetup.Header.AddParagraph("Header"), "Shown ", WordImageTextWrapping.InFrontOfText, height: 22D);
        WordParagraph inner = box.Paragraphs.Single();
        inner.AddHyperLink("SHOWNLINK", new Uri("https://example.com/shown"));
        W.Paragraph hidden = new W.Paragraph(new W.Run(new W.Text("Overflow ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
        box.Content!.Append(hidden);
        new WordParagraph(document, hidden).AddHyperLink("OVERFLOWLINK", new Uri("https://example.com/overflow"));

        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var links = pdf.GetPage(1).GetAnnotations().Where(annotation => annotation.Type == UglyToad.PdfPig.Annotations.AnnotationType.Link).ToArray();
        var link = Assert.Single(links);
        Assert.True(link.Rectangle.Bottom >= 792D - 24D - 22D - 0.01D, $"Link bottom {link.Rectangle.Bottom} is below the frame.");
    }

    // Paint the frame cannot reproduce keeps the whole header on its simple layout; a solid fill still places the box.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_HeaderTextBoxPaintDecidesThePlacedRouteBeforeLayout(bool gradient) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordTextBox box = AddPlacedHeaderBox(document.PageSetup.Header.AddParagraph("Header"), "PAINTED", WordImageTextWrapping.InFrontOfText);
        var properties = box.DrawingShapeProperties!;
        properties.RemoveAllChildren<A.SolidFill>();
        properties.RemoveAllChildren<A.NoFill>();
        OpenXmlElementFill(properties, gradient);

        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Page page = pdf.GetPage(1);
        Assert.Contains("PAINTED", string.Concat(page.Letters.Select(letter => letter.Value)), StringComparison.Ordinal);
        if (gradient) {
            // The simple header layout draws the box text without the bordered panel of the rich flowed route.
            Assert.DoesNotContain(page.Paths, path => path.IsStroked);
        } else {
            Assert.InRange(page.GetWords().Single(word => word.Text == "PAINTED").Letters[0].StartBaseLine.X, 300D, 312D);
        }
    }

    private static void OpenXmlElementFill(DocumentFormat.OpenXml.Office2010.Word.DrawingShape.ShapeProperties properties, bool gradient) {
        DocumentFormat.OpenXml.OpenXmlElement fill = gradient
            ? new A.GradientFill(new A.GradientStopList(
                new A.GradientStop(new A.RgbColorModelHex { Val = "FFFFFF" }) { Position = 0 },
                new A.GradientStop(new A.RgbColorModelHex { Val = "FFFF00" }) { Position = 100000 }), new A.LinearGradientFill { Angle = 0 })
            : new A.SolidFill(new A.RgbColorModelHex { Val = "FFFF00" });
        properties.InsertAfter(fill, properties.GetFirstChild<A.PresetGeometry>());
    }

    // The paragraph that anchors a placed box keeps its own shading.
    [Fact]
    public void SaveAsPdf_HeaderTextBoxAnchorParagraphKeepsItsShading() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body");
        WordParagraph anchorParagraph = document.PageSetup.Header.AddParagraph("Shaded header");
        anchorParagraph.ShadingFillColor = global::OfficeIMO.OfficeColor.Parse("00FF00");
        AddPlacedHeaderBox(anchorParagraph, "Box", WordImageTextWrapping.InFrontOfText);

        var result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Page page = pdf.GetPage(1);
        Assert.Contains(page.Paths, path => path.IsFilled && path.FillColor.ToRGBValues() == (0D, 1D, 0D));
        Assert.InRange(page.GetWords().Single(word => word.Text == "Box").Letters[0].StartBaseLine.X, 300D, 312D);
    }

    private static byte[] RenderWithFont(WordDocument document, byte[] font) {
        var pdfOptions = new OfficeIMO.Pdf.PdfCreateOptions();
        pdfOptions.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName, font));
        var options = new WordPdfExportOptions { IncludePageNumbers = false, PdfOptions = pdfOptions };
        return document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    }
}

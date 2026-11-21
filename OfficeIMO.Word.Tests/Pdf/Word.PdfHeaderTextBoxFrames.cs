using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfWord = UglyToad.PdfPig.Content.Word;

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

    private static byte[] RenderWithFont(WordDocument document, byte[] font) {
        var pdfOptions = new OfficeIMO.Pdf.PdfCreateOptions();
        pdfOptions.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName, font));
        var options = new WordPdfExportOptions { IncludePageNumbers = false, PdfOptions = pdfOptions };
        return document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    }
}

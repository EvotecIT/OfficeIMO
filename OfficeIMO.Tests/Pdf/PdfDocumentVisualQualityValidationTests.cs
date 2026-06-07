using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void PdfDocument_DefaultTextStyleRejectsInvalidInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultTextStyle((Action<PdfTextStyleCompose>)null!));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultTextStyle((PdfTextStyle)null!));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().Theme(null!));
        Assert.Throws<ArgumentNullException>(() => new PdfOptions().ApplyTheme(null!));

        var fontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create().DefaultTextStyle(style => style.Font((PdfStandardFont)99)));

        Assert.Equal("font", fontException.ParamName);
        Assert.Contains("PDF default font must be one of the supported standard PDF fonts.", fontException.Message, StringComparison.Ordinal);

        var fontSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create().DefaultTextStyle(style => style.FontSize(double.NaN)));

        Assert.Equal("size", fontSizeException.ParamName);

        var textStyleFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfTextStyle { Font = (PdfStandardFont)99 });

        Assert.Equal("Font", textStyleFontException.ParamName);
        Assert.Contains("PDF text style font must be one of the supported standard PDF fonts.", textStyleFontException.Message, StringComparison.Ordinal);

        var textStyleFontSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfTextStyle { FontSize = 0 });

        Assert.Equal("FontSize", textStyleFontSizeException.ParamName);

        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultHeadingStyle(1, null!));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().DefaultHeadingStyle(4, new PdfHeadingStyle()));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultPanelStyle(null!));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultHorizontalRuleStyle(null!));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultDrawingStyle(null!));

        var headingSizeException = Assert.Throws<ArgumentException>(() =>
            new PdfHeadingStyle { FontSize = 0 });

        Assert.Equal("FontSize", headingSizeException.ParamName);

        var headingSpacingException = Assert.Throws<ArgumentException>(() =>
            new PdfHeadingStyle { SpacingAfter = -1 });

        Assert.Equal("SpacingAfter", headingSpacingException.ParamName);
    }

    [Fact]
    public void Heading_RejectsUnsupportedAlignmentBeforeRendering() {
        var invalidAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H1("Invalid heading", (PdfAlign)99));

        Assert.Contains("Heading alignment must be Left, Center, or Right.", invalidAlignException.Message, StringComparison.Ordinal);

        var justifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H2("Unsupported heading", PdfAlign.Justify));

        Assert.Contains("Heading alignment must be Left, Center, or Right.", justifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Lists_RejectUnsupportedAlignmentBeforeRendering() {
        var invalidBulletAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Bullets(new[] { "Invalid bullet" }, (PdfAlign)99));

        Assert.Contains("Bullet list alignment must be Left, Center, or Right.", invalidBulletAlignException.Message, StringComparison.Ordinal);

        var bulletJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Bullets(new[] { "Unsupported bullet" }, PdfAlign.Justify));

        Assert.Contains("Bullet list alignment must be Left, Center, or Right.", bulletJustifyException.Message, StringComparison.Ordinal);

        var invalidNumberedAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Numbered(new[] { "Invalid numbered" }, (PdfAlign)99));

        Assert.Contains("Numbered list alignment must be Left, Center, or Right.", invalidNumberedAlignException.Message, StringComparison.Ordinal);

        var numberedJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Numbered(new[] { "Unsupported numbered" }, PdfAlign.Justify));

        Assert.Contains("Numbered list alignment must be Left, Center, or Right.", numberedJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageShapeAndDrawingBlocks_RejectUnsupportedAlignmentBeforeRendering() {
        byte[] png = CreateMinimalRgbPng();

        var invalidImageAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Image(png, 24, 24, (PdfAlign)99));

        Assert.Contains("Image alignment must be Left, Center, or Right.", invalidImageAlignException.Message, StringComparison.Ordinal);

        var imageJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Image(png, 24, 24, PdfAlign.Justify));

        Assert.Contains("Image alignment must be Left, Center, or Right.", imageJustifyException.Message, StringComparison.Ordinal);

        var shape = OfficeShape.Rectangle(24, 12);

        var invalidShapeAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Shape(shape, (PdfAlign)99));

        Assert.Contains("Shape alignment must be Left, Center, or Right.", invalidShapeAlignException.Message, StringComparison.Ordinal);

        var shapeJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Shape(shape, PdfAlign.Justify));

        Assert.Contains("Shape alignment must be Left, Center, or Right.", shapeJustifyException.Message, StringComparison.Ordinal);

        var drawing = new OfficeDrawing(24, 12)
            .AddShape(OfficeShape.Rectangle(24, 12), 0, 0);

        var invalidDrawingAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Drawing(drawing, (PdfAlign)99));

        Assert.Contains("Drawing alignment must be Left, Center, or Right.", invalidDrawingAlignException.Message, StringComparison.Ordinal);

        var drawingJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Drawing(drawing, PdfAlign.Justify));

        Assert.Contains("Drawing alignment must be Left, Center, or Right.", drawingJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphAndPanelBlocks_RejectInvalidAlignmentModelState() {
        var paragraphAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.Text("Invalid paragraph"), (PdfAlign)99));

        Assert.Contains("Paragraph alignment must be Left, Center, Right, or Justify.", paragraphAlignException.Message, StringComparison.Ordinal);

        byte[] justifiedParagraph = PdfDocument.Create()
            .Paragraph(p => p.Text("Justified paragraph alignment remains supported for report text."), PdfAlign.Justify)
            .ToBytes();

        Assert.NotEmpty(justifiedParagraph);

        var panelParagraphAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().PanelParagraph(p => p.Text("Invalid panel text"), new PanelStyle(), (PdfAlign)99));

        Assert.Contains("Panel paragraph alignment must be Left, Center, Right, or Justify.", panelParagraphAlignException.Message, StringComparison.Ordinal);

        byte[] justifiedPanelParagraph = PdfDocument.Create()
            .PanelParagraph(p => p.Text("Justified panel text remains supported inside the panel box."), new PanelStyle(), PdfAlign.Justify)
            .ToBytes();

        Assert.NotEmpty(justifiedPanelParagraph);

        var invalidPanelBoxAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .PanelParagraph(p => p.Text("Invalid panel box"), new PanelStyle {
                    MaxWidth = 120,
                    Align = (PdfAlign)99
                })
                .ToBytes());

        Assert.Contains("Panel box alignment must be Left, Center, or Right.", invalidPanelBoxAlignException.Message, StringComparison.Ordinal);

        var panelBoxJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .PanelParagraph(p => p.Text("Unsupported panel box"), new PanelStyle {
                    MaxWidth = 120,
                    Align = PdfAlign.Justify
                })
                .ToBytes());

        Assert.Contains("Panel box alignment must be Left, Center, or Right.", panelBoxJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MutableAlignmentProperties_RejectUnsupportedValuesOnAssignment() {
        var headerAlignException = Assert.Throws<ArgumentException>(() =>
            new PdfOptions {
                HeaderAlign = (PdfAlign)99
            });

        Assert.Contains("PDF header alignment must be Left, Center, or Right.", headerAlignException.Message, StringComparison.Ordinal);

        var headerJustifyException = Assert.Throws<ArgumentException>(() =>
            new PdfOptions {
                HeaderAlign = PdfAlign.Justify
            });

        Assert.Contains("PDF header alignment must be Left, Center, or Right.", headerJustifyException.Message, StringComparison.Ordinal);

        var footerAlignException = Assert.Throws<ArgumentException>(() =>
            new PdfOptions {
                FooterAlign = (PdfAlign)99
            });

        Assert.Contains("PDF footer alignment must be Left, Center, or Right.", footerAlignException.Message, StringComparison.Ordinal);

        var footerJustifyException = Assert.Throws<ArgumentException>(() =>
            new PdfOptions {
                FooterAlign = PdfAlign.Justify
            });

        Assert.Contains("PDF footer alignment must be Left, Center, or Right.", footerJustifyException.Message, StringComparison.Ordinal);

        var panelAlignException = Assert.Throws<ArgumentException>(() =>
            new PanelStyle {
                Align = (PdfAlign)99
            });

        Assert.Contains("Panel box alignment must be Left, Center, or Right.", panelAlignException.Message, StringComparison.Ordinal);

        var panelJustifyException = Assert.Throws<ArgumentException>(() =>
            new PanelStyle {
                Align = PdfAlign.Justify
            });

        Assert.Contains("Panel box alignment must be Left, Center, or Right.", panelJustifyException.Message, StringComparison.Ordinal);

        var captionStyle = TableStyles.Minimal();
        var captionAlignException = Assert.Throws<ArgumentException>(() =>
            captionStyle.CaptionAlign = (PdfAlign)99);

        Assert.Contains("Table caption alignment must be Left, Center, or Right.", captionAlignException.Message, StringComparison.Ordinal);

        var captionJustifyException = Assert.Throws<ArgumentException>(() =>
            captionStyle.CaptionAlign = PdfAlign.Justify);

        Assert.Contains("Table caption alignment must be Left, Center, or Right.", captionJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void LinkAnnotations_RejectInvalidUriModelStateBeforeRendering() {
        Assert.Throws<ArgumentNullException>(() =>
            PdfDocument.Create().Paragraph(p => p.Link("OfficeIMO", null!)));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.Link("OfficeIMO", "bad\u0001uri")));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.Link("", "https://evotec.xyz")));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.Link("OfficeIMO", "https://evotec.xyz", contents: " ")));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H1("Linked heading", linkUri: "bad\u0001uri"));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H1("Linked heading", linkUri: "https://evotec.xyz", linkContents: " "));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H1("Plain heading", linkContents: "metadata without link"));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H1("Conflicting heading link", linkUri: "https://evotec.xyz", linkDestinationName: "Intro"));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H1("Bookmark heading link", linkDestinationName: " ", linkContents: "metadata"));

        byte[] png = CreateMinimalRgbPng();
        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Image(png, 24, 24, linkUri: "bad\u0001uri"));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Image(png, 24, 24, linkUri: "https://evotec.xyz", linkContents: " "));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Image(png, 24, 24, linkContents: "metadata without link"));

        var shape = OfficeShape.Rectangle(24, 12);
        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Shape(shape, linkUri: "bad\u0001uri"));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Shape(shape, linkUri: "https://evotec.xyz", linkContents: " "));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Shape(shape, linkContents: "metadata without link"));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Rectangle(24, 12, linkUri: "bad\u0001uri"));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Rectangle(24, 12, linkUri: "https://evotec.xyz", linkContents: " "));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Rectangle(24, 12, linkContents: "metadata without link"));

        var drawing = new OfficeDrawing(24, 12)
            .AddShape(OfficeShape.Rectangle(24, 12), 0, 0);
        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Drawing(drawing, linkUri: "bad\u0001uri"));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Drawing(drawing, linkUri: "https://evotec.xyz", linkContents: " "));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Drawing(drawing, linkContents: "metadata without link"));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().TableWithLinks(
                new[] { new[] { "Name" }, new[] { "OfficeIMO" } },
                new Dictionary<(int Row, int Col), string> {
                    [(1, 0)] = "bad\u0001uri"
                }));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create().TableWithLinks(
                new[] { new[] { "Name" }, new[] { "OfficeIMO" } },
                new Dictionary<(int Row, int Col), string> {
                    [(-1, 0)] = "https://evotec.xyz"
                }));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create().TableWithLinks(
                new[] { new[] { "Name" }, new[] { "OfficeIMO" } },
                new Dictionary<(int Row, int Col), string> {
                    [(2, 0)] = "https://evotec.xyz"
                }));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create().TableWithLinks(
                new[] { new[] { "Name", "Url" }, new[] { "OfficeIMO" } },
                new Dictionary<(int Row, int Col), string> {
                    [(1, 1)] = "https://evotec.xyz"
                }));
    }


}

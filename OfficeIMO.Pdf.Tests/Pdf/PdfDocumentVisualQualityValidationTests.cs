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
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultTextStyle((Action<PdfTextStyleBuilder>)null!));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultTextStyle((PdfTextStyle)null!));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().Theme(null!));
        Assert.Throws<ArgumentNullException>(() => new PdfCreateOptions().ApplyTheme(null!));

        var fontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDocument.Create().DefaultTextStyle(style => style.Font((PdfStandardFont)99)));

        Assert.Equal("font", fontException.ParamName);
        Assert.Contains("PDF default font must be one of the supported standard PDF fonts.", fontException.Message, StringComparison.Ordinal);





        var textStyleFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfTextStyle { Font = (PdfStandardFont)99 });

        Assert.Equal("Font", textStyleFontException.ParamName);
        Assert.Contains("PDF text style font must be one of the supported standard PDF fonts.", textStyleFontException.Message, StringComparison.Ordinal);

        var textStyleFontSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfTextStyle { FontSize = global::OfficeIMO.OfficeLength.Points(0) });

        Assert.Equal("FontSize", textStyleFontSizeException.ParamName);

        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultHeadingStyle(1, null!));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().DefaultHeadingStyle(4, new PdfHeadingStyle()));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultPanelStyle(null!));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultHorizontalRuleStyle(null!));
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultDrawingStyle(null!));

        var headingSizeException = Assert.Throws<ArgumentException>(() =>
            new PdfHeadingStyle { FontSize = global::OfficeIMO.OfficeLength.Points(0) });

        Assert.Equal("FontSize", headingSizeException.ParamName);

        var headingSpacingException = Assert.Throws<ArgumentException>(() =>
            new PdfHeadingStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(-1) });

        Assert.Equal("SpacingAfter", headingSpacingException.ParamName);
    }

    [Fact]
    public void Heading_RejectsUnsupportedAlignmentBeforeRendering() {
        var invalidAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H1("Invalid heading", (OfficeAlignment)99));

        Assert.Contains("Heading alignment must be Left, Center, Right, Start, or End.", invalidAlignException.Message, StringComparison.Ordinal);

        var justifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().H2("Unsupported heading", OfficeAlignment.Justify));

        Assert.Contains("Heading alignment must be Left, Center, Right, Start, or End.", justifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Lists_RejectUnsupportedAlignmentBeforeRendering() {
        var invalidBulletAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Bullets(new[] { "Invalid bullet" }, (OfficeAlignment)99));

        Assert.Contains("Bullet list alignment must be Left, Center, Right, Start, or End.", invalidBulletAlignException.Message, StringComparison.Ordinal);

        var bulletJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Bullets(new[] { "Unsupported bullet" }, OfficeAlignment.Justify));

        Assert.Contains("Bullet list alignment must be Left, Center, Right, Start, or End.", bulletJustifyException.Message, StringComparison.Ordinal);

        var invalidNumberedAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Numbered(new[] { "Invalid numbered" }, (OfficeAlignment)99));

        Assert.Contains("Numbered list alignment must be Left, Center, Right, Start, or End.", invalidNumberedAlignException.Message, StringComparison.Ordinal);

        var numberedJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Numbered(new[] { "Unsupported numbered" }, OfficeAlignment.Justify));

        Assert.Contains("Numbered list alignment must be Left, Center, Right, Start, or End.", numberedJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageShapeAndDrawingBlocks_RejectUnsupportedAlignmentBeforeRendering() {
        byte[] png = CreateMinimalRgbPng();

        var invalidImageAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Image(png, 24, 24, (OfficeAlignment)99));

        Assert.Contains("Image alignment must be Left, Center, Right, Start, or End.", invalidImageAlignException.Message, StringComparison.Ordinal);

        var imageJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Image(png, 24, 24, OfficeAlignment.Justify));

        Assert.Contains("Image alignment must be Left, Center, Right, Start, or End.", imageJustifyException.Message, StringComparison.Ordinal);

        var shape = OfficeShape.Rectangle(24, 12);

        var invalidShapeAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Shape(shape, (OfficeAlignment)99));

        Assert.Contains("Shape alignment must be Left, Center, Right, Start, or End.", invalidShapeAlignException.Message, StringComparison.Ordinal);

        var shapeJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Shape(shape, OfficeAlignment.Justify));

        Assert.Contains("Shape alignment must be Left, Center, Right, Start, or End.", shapeJustifyException.Message, StringComparison.Ordinal);

        var drawing = new OfficeDrawing(24, 12)
            .AddShape(OfficeShape.Rectangle(24, 12), 0, 0);

        var invalidDrawingAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Drawing(drawing, (OfficeAlignment)99));

        Assert.Contains("Drawing alignment must be Left, Center, Right, Start, or End.", invalidDrawingAlignException.Message, StringComparison.Ordinal);

        var drawingJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Drawing(drawing, OfficeAlignment.Justify));

        Assert.Contains("Drawing alignment must be Left, Center, Right, Start, or End.", drawingJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphAndPanelBlocks_RejectInvalidAlignmentModelState() {
        var paragraphAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.AddText("Invalid paragraph"), (OfficeAlignment)99));

        Assert.Contains("Paragraph alignment must be Left, Center, Right, Justify, Start, or End.", paragraphAlignException.Message, StringComparison.Ordinal);

        byte[] justifiedParagraph = PdfDocument.Create()
            .Paragraph(p => p.AddText("Justified paragraph alignment remains supported for report text."), OfficeAlignment.Justify)
            .ToBytes();

        Assert.NotEmpty(justifiedParagraph);

        var panelParagraphAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().PanelParagraph(p => p.AddText("Invalid panel text"), new PdfPanelStyle(), (OfficeAlignment)99));

        Assert.Contains("Paragraph alignment must be Left, Center, Right, Justify, Start, or End.", panelParagraphAlignException.Message, StringComparison.Ordinal);

        byte[] justifiedPanelParagraph = PdfDocument.Create()
            .PanelParagraph(p => p.AddText("Justified panel text remains supported inside the panel box."), new PdfPanelStyle(), OfficeAlignment.Justify)
            .ToBytes();

        Assert.NotEmpty(justifiedPanelParagraph);

        var invalidPanelBoxAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .PanelParagraph(p => p.AddText("Invalid panel box"), new PdfPanelStyle {
                    MaxWidth = global::OfficeIMO.OfficeLength.Points(120),
                    Align = (OfficeAlignment)99
                })
                .ToBytes());

        Assert.Contains("Panel box alignment must be Left, Center, Right, Start, or End.", invalidPanelBoxAlignException.Message, StringComparison.Ordinal);

        var panelBoxJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .PanelParagraph(p => p.AddText("Unsupported panel box"), new PdfPanelStyle {
                    MaxWidth = global::OfficeIMO.OfficeLength.Points(120),
                    Align = OfficeAlignment.Justify
                })
                .ToBytes());

        Assert.Contains("Panel box alignment must be Left, Center, Right, Start, or End.", panelBoxJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MutableAlignmentProperties_RejectUnsupportedValuesOnAssignment() {
        var headerAlignException = Assert.Throws<ArgumentException>(() =>
            new PdfCreateOptions {
                PageSetup = { Header = { Alignment = (OfficeAlignment)99 } }
            });

        Assert.Contains("PDF header alignment must be Left, Center, Right, Start, or End.", headerAlignException.Message, StringComparison.Ordinal);

        var headerJustifyException = Assert.Throws<ArgumentException>(() =>
            new PdfCreateOptions {
                PageSetup = { Header = { Alignment = OfficeAlignment.Justify } }
            });

        Assert.Contains("PDF header alignment must be Left, Center, Right, Start, or End.", headerJustifyException.Message, StringComparison.Ordinal);

        var footerAlignException = Assert.Throws<ArgumentException>(() =>
            new PdfCreateOptions {
                PageSetup = { Footer = { Alignment = (OfficeAlignment)99 } }
            });

        Assert.Contains("PDF footer alignment must be Left, Center, Right, Start, or End.", footerAlignException.Message, StringComparison.Ordinal);

        var footerJustifyException = Assert.Throws<ArgumentException>(() =>
            new PdfCreateOptions {
                PageSetup = { Footer = { Alignment = OfficeAlignment.Justify } }
            });

        Assert.Contains("PDF footer alignment must be Left, Center, Right, Start, or End.", footerJustifyException.Message, StringComparison.Ordinal);

        var panelAlignException = Assert.Throws<ArgumentException>(() =>
            new PdfPanelStyle {
                Align = (OfficeAlignment)99
            });

        Assert.Contains("Panel box alignment must be Left, Center, Right, Start, or End.", panelAlignException.Message, StringComparison.Ordinal);

        var panelJustifyException = Assert.Throws<ArgumentException>(() =>
            new PdfPanelStyle {
                Align = OfficeAlignment.Justify
            });

        Assert.Contains("Panel box alignment must be Left, Center, Right, Start, or End.", panelJustifyException.Message, StringComparison.Ordinal);

        var captionStyle = PdfTableStyles.Minimal();
        var captionAlignException = Assert.Throws<ArgumentException>(() =>
            captionStyle.CaptionAlign = (OfficeAlignment)99);

        Assert.Contains("Table caption alignment must be Left, Center, Right, Start, or End.", captionAlignException.Message, StringComparison.Ordinal);

        var captionJustifyException = Assert.Throws<ArgumentException>(() =>
            captionStyle.CaptionAlign = OfficeAlignment.Justify);

        Assert.Contains("Table caption alignment must be Left, Center, Right, Start, or End.", captionJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void LinkAnnotations_RejectInvalidUriModelStateBeforeRendering() {
        Assert.Throws<ArgumentNullException>(() =>
            PdfDocument.Create().Paragraph(p => p.AddLink("OfficeIMO", null!)));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.AddLink("OfficeIMO", "bad\u0001uri")));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.AddLink("", "https://evotec.xyz")));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.AddLink("OfficeIMO", "https://evotec.xyz", contents: " ")));

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
    }


}

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
    public void TextWatermark_RendersBehindContentWithOpacityAndRotation() {
        byte[] bytes = PdfDocument.Create()
            .Watermark("DRAFT", fontSize: 48, color: PdfColor.FromRgb(120, 130, 150), opacity: 0.18, rotationAngle: -45)
            .H1("Watermark proof")
            .Paragraph(p => p.Text("Body content stays readable above the watermark."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("Watermark proof", text);
        Assert.Contains("DRAFT", text);
        Assert.Contains("/ExtGState", raw);
        Assert.Contains("/ca 0.18", raw, StringComparison.Ordinal);
        Assert.Contains("0.707 -0.707 0.707 0.707", raw, StringComparison.Ordinal);
        Assert.Contains("<4452414654> Tj", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void TextWatermark_CanBeScopedAndClearedPerComposedPage() {
        byte[] bytes = PdfDocument.Create()
            .Watermark("GLOBAL", fontSize: 30, opacity: 0.16)
            .Page(page => {
                page.Watermark("SECTION", fontSize: 30, opacity: 0.16);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Section page"))));
            })
            .Page(page => {
                page.Watermark((PdfTextWatermark?)null);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Clean page"))));
            })
            .ToBytes();

        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("SECTION", text);
        Assert.DoesNotContain("GLOBAL", text);
    }

    [Fact]
    public void TextWatermark_ValidatesAndClonesOptions() {
        var options = new PdfOptions {
            TextWatermark = new PdfTextWatermark("Original") {
                Opacity = 0.2,
                RotationAngle = -20,
                FontSize = 40
            }
        };

        PdfTextWatermark snapshot = options.TextWatermark!;
        snapshot.Text = "Changed";

        Assert.Equal("Original", options.TextWatermark!.Text);
        Assert.Throws<ArgumentException>(() => new PdfTextWatermark(""));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextWatermark("Bad") { Opacity = 1.5 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextWatermark("Bad") { RotationAngle = double.NaN });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextWatermark("Bad") { Font = (PdfStandardFont)99 });
    }

    [Fact]
    public void ImageWatermark_RendersBehindContentWithOpacityAndRotation() {
        byte[] bytes = PdfDocument.Create()
            .ImageWatermark(CreateMinimalRgbPng(), width: 120, height: 60, opacity: 0.2, rotationAngle: 30)
            .H1("Image watermark proof")
            .Paragraph(p => p.Text("Body content stays readable above the image watermark."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string stream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        int imageDraw = stream.IndexOf("/Im", StringComparison.Ordinal);
        int firstTextBlock = stream.IndexOf("BT", StringComparison.Ordinal);

        Assert.Contains("/Subtype /Image", raw, StringComparison.Ordinal);
        Assert.Contains("/ExtGState", raw, StringComparison.Ordinal);
        Assert.Contains("/ca 0.2", raw, StringComparison.Ordinal);
        Assert.True(imageDraw >= 0, "Expected the image watermark XObject to be drawn.");
        Assert.True(firstTextBlock > imageDraw, "Expected body text to be emitted after the image watermark draw command.");
    }

    [Fact]
    public void ImageWatermark_CanBeScopedAndClearedPerComposedPage() {
        byte[] image = CreateMinimalRgbPng();
        byte[] bytes = PdfDocument.Create()
            .ImageWatermark(image, width: 80, height: 80, opacity: 0.18)
            .Page(page => {
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Marked page"))));
            })
            .Page(page => {
                page.ImageWatermark((PdfImageWatermark?)null);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Clean page"))));
            })
            .ToBytes();

        string firstPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        string secondPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 2));

        Assert.Contains("/Im", firstPageStream, StringComparison.Ordinal);
        Assert.DoesNotContain("/Im", secondPageStream, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageWatermark_FirstAndEvenVariantsUseMatchingPageVariant() {
        byte[] image = CreateMinimalRgbPng();
        var options = new PdfOptions {
            DifferentFirstPageHeaderFooter = true,
            DifferentOddAndEvenPagesHeaderFooter = true,
            ImageWatermark = new PdfImageWatermark(image, width: 20, height: 20) {
                Opacity = 0.18
            },
            FirstPageImageWatermark = new PdfImageWatermark(image, width: 21, height: 21) {
                Opacity = 0.18
            },
            EvenPageImageWatermark = new PdfImageWatermark(image, width: 22, height: 22) {
                Opacity = 0.18
            }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Page one body."))
            .PageBreak()
            .Paragraph(p => p.Text("Page two body."))
            .PageBreak()
            .Paragraph(p => p.Text("Page three body."))
            .ToBytes();

        string firstPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        string secondPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 2));
        string thirdPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 3));

        Assert.Contains("21 0 0 21", firstPageStream, StringComparison.Ordinal);
        Assert.Contains("22 0 0 22", secondPageStream, StringComparison.Ordinal);
        Assert.Contains("20 0 0 20", thirdPageStream, StringComparison.Ordinal);
    }

    [Fact]
    public void WatermarkVariants_DoNotEnableHeaderFooterVariantsByThemselves() {
        var options = new PdfOptions {
            ShowHeader = true,
            HeaderFormat = "DEFAULT HEADER",
            FirstPageTextWatermark = new PdfTextWatermark("FIRST") {
                Opacity = 0.18,
                FontSize = 30
            }
        };

        Assert.False(options.DifferentFirstPageHeaderFooter);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("First page body."))
            .PageBreak()
            .Paragraph(p => p.Text("Second page body."))
            .ToBytes();

        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("FIRST", text);
        Assert.Equal(2, Regex.Matches(text, "DEFAULT HEADER").Count);
    }

    [Fact]
    public void WatermarkVariants_FallBackToGlobalWatermarksWhenUnset() {
        byte[] image = CreateMinimalRgbPng();
        var options = new PdfOptions {
            DifferentFirstPageHeaderFooter = true,
            DifferentOddAndEvenPagesHeaderFooter = true,
            TextWatermark = new PdfTextWatermark("GLOBAL") {
                Opacity = 0.18,
                FontSize = 30
            },
            ImageWatermark = new PdfImageWatermark(image, width: 20, height: 20) {
                Opacity = 0.18
            }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Page one body."))
            .PageBreak()
            .Paragraph(p => p.Text("Page two body."))
            .PageBreak()
            .Paragraph(p => p.Text("Page three body."))
            .ToBytes();

        string text = PdfReadDocument.Open(bytes).ExtractText();
        string firstPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        string secondPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 2));
        string thirdPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 3));

        Assert.Equal(3, Regex.Matches(text, "GLOBAL").Count);
        Assert.Contains("/Im", firstPageStream, StringComparison.Ordinal);
        Assert.Contains("/Im", secondPageStream, StringComparison.Ordinal);
        Assert.Contains("/Im", thirdPageStream, StringComparison.Ordinal);
    }

    [Fact]
    public void WatermarkVariants_CanSuppressInheritedFirstPageWatermarks() {
        byte[] image = CreateMinimalRgbPng();
        byte[] bytes = PdfDocument.Create()
            .Watermark("GLOBAL", fontSize: 30, opacity: 0.18)
            .ImageWatermark(image, width: 20, height: 20, opacity: 0.18)
            .Page(page => {
                page.SuppressFirstPageWatermark();
                page.Content(content => content.Column(column => {
                    column.Item().Paragraph(p => p.Text("First page body."));
                    column.Item().PageBreak();
                    column.Item().Paragraph(p => p.Text("Second page body."));
                }));
            })
            .ToBytes();

        PdfReadDocument readDocument = PdfReadDocument.Open(bytes);
        string firstPageText = readDocument.Pages[0].ExtractText();
        string secondPageText = readDocument.Pages[1].ExtractText();
        string firstPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 1));
        string secondPageStream = Assert.Single(GetPageContentStreams(bytes, pageNumber: 2));

        Assert.DoesNotContain("GLOBAL", firstPageText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("GLOBAL", secondPageText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("/Im", firstPageStream, StringComparison.Ordinal);
        Assert.Contains("/Im", secondPageStream, StringComparison.Ordinal);
    }

    [Fact]
    public void WatermarkSuppression_DoesNotEnableHeaderFooterVariantsByItself() {
        var options = new PdfOptions {
            ShowHeader = true,
            HeaderFormat = "DEFAULT HEADER",
            TextWatermark = new PdfTextWatermark("GLOBAL") {
                Opacity = 0.18,
                FontSize = 30
            }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Page(page => {
                page.SuppressFirstPageWatermark();
                page.Content(content => content.Column(column => {
                    column.Item().Paragraph(p => p.Text("First page body."));
                    column.Item().PageBreak();
                    column.Item().Paragraph(p => p.Text("Second page body."));
                }));
            })
            .ToBytes();

        PdfReadDocument readDocument = PdfReadDocument.Open(bytes);
        string text = readDocument.ExtractText();
        string firstPageText = readDocument.Pages[0].ExtractText();

        Assert.False(options.DifferentFirstPageHeaderFooter);
        Assert.Equal(2, Regex.Matches(text, "DEFAULT HEADER").Count);
        Assert.DoesNotContain("GLOBAL", firstPageText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ImageWatermark_ValidatesAndClonesOptions() {
        byte[] image = CreateMinimalRgbPng();
        var options = new PdfOptions {
            ImageWatermark = new PdfImageWatermark(image, width: 20, height: 10) {
                Opacity = 0.25,
                RotationAngle = -15
            }
        };

        PdfImageWatermark snapshot = options.ImageWatermark!;
        snapshot.Width = 120;

        Assert.Equal(20, options.ImageWatermark!.Width);
        Assert.Throws<ArgumentException>(() => new PdfImageWatermark(Array.Empty<byte>(), 20, 10));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageWatermark(image, 0, 10));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageWatermark(image, 20, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageWatermark(image, 20, 10) { Opacity = 1.5 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageWatermark(image, 20, 10) { RotationAngle = double.PositiveInfinity });
    }


}

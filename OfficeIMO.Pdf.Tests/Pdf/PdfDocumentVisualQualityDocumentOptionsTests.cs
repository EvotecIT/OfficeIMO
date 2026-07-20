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
    public void Options_RejectInvalidPageGeometryAndTypography() {
        var widthException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 0
                })
                .Paragraph(p => p.Text("Invalid page width"))
                .ToBytes());

        Assert.Contains("PDF page width must be a positive finite value.", widthException.Message, StringComparison.Ordinal);

        var marginException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    MarginLeft = -1
                })
                .Paragraph(p => p.Text("Invalid margin"))
                .ToBytes());

        Assert.Contains("PDF left margin must be a non-negative finite value.", marginException.Message, StringComparison.Ordinal);

        var contentWidthException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 100,
                    MarginLeft = 50,
                    MarginRight = 50
                })
                .Paragraph(p => p.Text("No content width"))
                .ToBytes());

        Assert.Contains("PDF margins must leave a positive content width.", contentWidthException.Message, StringComparison.Ordinal);

        var contentHeightException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageHeight = 100,
                    MarginTop = 60,
                    MarginBottom = 40
                })
                .Paragraph(p => p.Text("No content height"))
                .ToBytes());

        Assert.Contains("PDF margins must leave a positive content height.", contentHeightException.Message, StringComparison.Ordinal);

        var fontException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    DefaultFontSize = double.NaN
                })
                .Paragraph(p => p.Text("Invalid font size"))
                .ToBytes());

        Assert.Contains("PDF default font size must be a positive finite value.", fontException.Message, StringComparison.Ordinal);

        var defaultFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfOptions {
                DefaultFont = (PdfStandardFont)99
            });

        Assert.Equal("DefaultFont", defaultFontException.ParamName);
        Assert.Contains("PDF default font must be one of the supported standard PDF fonts.", defaultFontException.Message, StringComparison.Ordinal);

        var headerException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    ShowHeader = true,
                    HeaderFormat = "Header",
                    HeaderFontSize = double.NaN
                })
                .Paragraph(p => p.Text("Invalid header font size"))
                .ToBytes());

        Assert.Contains("PDF header font size must be a positive finite value.", headerException.Message, StringComparison.Ordinal);

        var headerFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfOptions {
                HeaderFont = (PdfStandardFont)99
            });

        Assert.Equal("HeaderFont", headerFontException.ParamName);
        Assert.Contains("PDF header font must be one of the supported standard PDF fonts.", headerFontException.Message, StringComparison.Ordinal);

        var headerOffsetException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    ShowHeader = true,
                    HeaderFormat = "Header",
                    MarginTop = 20,
                    HeaderOffsetY = 21
                })
                .Paragraph(p => p.Text("Header above page"))
                .ToBytes());

        Assert.Contains("PDF header offset must not exceed the top margin when header content is enabled.", headerOffsetException.Message, StringComparison.Ordinal);

        var headerFormatException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    ShowHeader = true,
                    HeaderFormat = null!
                })
                .Paragraph(p => p.Text("Invalid header format"))
                .ToBytes());

        Assert.Contains("PDF header format cannot be null.", headerFormatException.Message, StringComparison.Ordinal);

        var headerAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    ShowHeader = true,
                    HeaderFormat = "Header",
                    HeaderAlign = (PdfAlign)99
                })
                .Paragraph(p => p.Text("Invalid header alignment"))
                .ToBytes());

        Assert.Contains("PDF header alignment must be Left, Center, or Right.", headerAlignException.Message, StringComparison.Ordinal);

        var headerJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    ShowHeader = true,
                    HeaderFormat = "Header",
                    HeaderAlign = PdfAlign.Justify
                })
                .Paragraph(p => p.Text("Unsupported header alignment"))
                .ToBytes());

        Assert.Contains("PDF header alignment must be Left, Center, or Right.", headerJustifyException.Message, StringComparison.Ordinal);

        var footerException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    ShowPageNumbers = true,
                    FooterFontSize = double.PositiveInfinity
                })
                .Paragraph(p => p.Text("Invalid footer font size"))
                .ToBytes());

        Assert.Contains("PDF footer font size must be a positive finite value.", footerException.Message, StringComparison.Ordinal);

        var footerFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfOptions {
                FooterFont = (PdfStandardFont)99
            });

        Assert.Equal("FooterFont", footerFontException.ParamName);
        Assert.Contains("PDF footer font must be one of the supported standard PDF fonts.", footerFontException.Message, StringComparison.Ordinal);

        var footerOffsetException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    ShowPageNumbers = true,
                    MarginBottom = 20,
                    FooterOffsetY = 21
                })
                .Paragraph(p => p.Text("Footer below page"))
                .ToBytes());

        Assert.Contains("PDF footer offset must not exceed the bottom margin when footer content is enabled.", footerOffsetException.Message, StringComparison.Ordinal);

        var footerAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    ShowPageNumbers = true,
                    FooterAlign = (PdfAlign)99
                })
                .Paragraph(p => p.Text("Invalid footer alignment"))
                .ToBytes());

        Assert.Contains("PDF footer alignment must be Left, Center, or Right.", footerAlignException.Message, StringComparison.Ordinal);

        var footerJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    ShowPageNumbers = true,
                    FooterAlign = PdfAlign.Justify
                })
                .Paragraph(p => p.Text("Unsupported footer alignment"))
                .ToBytes());

        Assert.Contains("PDF footer alignment must be Left, Center, or Right.", footerJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ComposePage_RejectsInvalidPageOptions() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Compose(compose =>
                    compose.Page(page => {
                        page.Size(200, 160);
                        page.Margin(left: 100, top: 20, right: 100, bottom: 20);
                        page.Content(content =>
                            content.Column(column =>
                                column.Item().Paragraph(p => p.Text("No content width"))));
                    }))
                .ToBytes());

        Assert.Contains("PDF margins must leave a positive content width.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DefaultOptions_UseProportionalHelveticaForPlainDocuments() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                ShowHeader = true,
                HeaderFormat = "HeaderMarker",
                ShowPageNumbers = true,
                FooterFormat = "FooterMarker"
            })
            .Paragraph(p => p.Text("BodyMarker uses the built-in default font."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/BaseFont /Helvetica", content);
        Assert.DoesNotContain("/BaseFont /Courier", content);
    }

    [Fact]
    public void StandardFontMapper_MapsOfficeFamiliesToDependencyFreePdfFonts() {
        Assert.True(PdfStandardFontMapper.TryMapFontFamily("Segoe UI, sans-serif", out PdfStandardFont sans));
        Assert.Equal(PdfStandardFont.Helvetica, sans);

        Assert.True(PdfStandardFontMapper.TryMapFontFamily("Unmapped Display Face, Georgia, serif", out PdfStandardFont fallbackList));
        Assert.Equal(PdfStandardFont.TimesRoman, fallbackList);

        Assert.True(PdfStandardFontMapper.TryMapFontFamily("\"Times New Roman\", serif", bold: true, italic: true, out PdfStandardFont serif));
        Assert.Equal(PdfStandardFont.TimesBoldItalic, serif);

        Assert.True(PdfStandardFontMapper.TryMapFontFamily("Consolas", bold: false, italic: true, out PdfStandardFont mono));
        Assert.Equal(PdfStandardFont.CourierOblique, mono);

        Assert.False(PdfStandardFontMapper.TryMapFontFamily("Unmapped Display Face", out PdfStandardFont fallback));
        Assert.Equal(PdfStandardFont.Helvetica, fallback);

        Assert.True(PdfStandardFontMapper.IsStandardPdfFamilyEquivalent("Helvetica, sans-serif", PdfStandardFont.Helvetica));
        Assert.True(PdfStandardFontMapper.IsStandardPdfFamilyEquivalent("serif", PdfStandardFont.TimesRoman));
        Assert.True(PdfStandardFontMapper.IsStandardPdfFamilyEquivalent("monospace", PdfStandardFont.Courier));
        Assert.False(PdfStandardFontMapper.IsStandardPdfFamilyEquivalent("Arial, sans-serif", PdfStandardFont.Helvetica));
        Assert.False(PdfStandardFontMapper.IsStandardPdfFamilyEquivalent("Times New Roman, serif", PdfStandardFont.TimesRoman));

        Assert.Equal(PdfStandardFont.TimesBold, PdfStandardFontMapper.GetStyledFont(PdfStandardFont.TimesItalic, bold: true, italic: false));
    }


}

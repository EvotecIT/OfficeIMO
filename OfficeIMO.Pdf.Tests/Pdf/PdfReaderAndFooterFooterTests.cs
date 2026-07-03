using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReaderAndFooterRegressionTests {

    [Fact]
    public void Footer_UsesConfiguredFooterFont() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.TimesRoman,
            ShowPageNumbers = true,
            FooterFormat = "Footer check",
            FooterFont = PdfStandardFont.HelveticaBold,
            FooterAlign = PdfAlign.Left
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Body text only."))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var footerLine = pdf.GetPage(1).Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderBy(group => group.Key)
            .First()
            .OrderBy(letter => letter.StartBaseLine.X)
            .ToList();

        string footerText = new string(string.Concat(footerLine.Select(letter => letter.Value)).Where(c => !char.IsWhiteSpace(c)).ToArray());
        Assert.Equal("Footercheck", footerText);
        Assert.Contains(footerLine, letter => letter.FontName != null && letter.FontName.Contains("Helvetica-Bold", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(footerLine, letter => letter.FontName != null && letter.FontName.Contains("Times", StringComparison.OrdinalIgnoreCase));

        string content = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/BaseFont /Helvetica-Bold", content);
    }

    [Fact]
    public void FooterFontResource_IsNotEmittedWhenFooterIsDisabled() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            FooterFont = PdfStandardFont.CourierBold,
            ShowPageNumbers = false
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Body text only."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/BaseFont /Helvetica", content);
        Assert.DoesNotContain("/BaseFont /Courier-Bold", content);
        Assert.DoesNotContain("/F5", content);
    }

    [Fact]
    public void FooterSegments_RenderWithoutShowPageNumbersFlag() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.TimesRoman,
            FooterFont = PdfStandardFont.HelveticaBold,
            FooterAlign = PdfAlign.Left,
            FooterSegments = new System.Collections.Generic.List<FooterSegment> {
                new FooterSegment(FooterSegmentKind.Text, "Direct footer"),
                new FooterSegment(FooterSegmentKind.Text, " "),
                new FooterSegment(FooterSegmentKind.PageNumber)
            }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Body text only."))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string pageText = new string(pdf.GetPage(1).Text.Where(c => !char.IsWhiteSpace(c)).ToArray());
        Assert.Contains("Directfooter1", pageText, StringComparison.OrdinalIgnoreCase);

        string content = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/BaseFont /Helvetica-Bold", content);
    }

    [Fact]
    public void FooterSegments_ValidateFooterPlacementWithoutShowPageNumbersFlag() {
        var options = new PdfOptions {
            MarginBottom = 20,
            FooterOffsetY = 21,
            FooterSegments = new System.Collections.Generic.List<FooterSegment> {
                new FooterSegment(FooterSegmentKind.Text, "Direct footer")
            }
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(options)
                .Paragraph(p => p.Text("Body text only."))
                .ToBytes());

        Assert.Contains("PDF footer offset must not exceed the bottom margin when footer content is enabled.", exception.Message, StringComparison.Ordinal);
    }

}

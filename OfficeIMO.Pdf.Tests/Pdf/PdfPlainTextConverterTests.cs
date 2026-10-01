using System;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPlainTextConverterTests {
    [Fact]
    public void StreamImportLeavesOwnershipAndPositionWithCallerAndEnforcesByteLimit() {
        using var input = new System.IO.MemoryStream(Encoding.UTF8.GetBytes("skip Stream text"));
        input.Position = 5;
        using var pdf = PdfPigDocument.Open(PdfPlainTextConverter.ToPdfDocumentResult(input).ToBytes());
        Assert.Contains("Stream text", pdf.GetPage(1).Text);
        Assert.Equal(5, input.Position); Assert.True(input.CanRead);
        Assert.Throws<System.IO.InvalidDataException>(() => PdfPlainTextConverter.ToPdfDocumentResult(input, maximumInputBytes: 2));
        Assert.Equal(5, input.Position);
    }
    [Theory]
    [InlineData("", 1)]
    [InlineData("   \n\n", 1)]
    [InlineData("\f", 2)]
    public void EmptyPagesRemainValidPdfPages(string text, int pages) {
        using var pdf = PdfPigDocument.Open(PdfPlainTextConverter.ToPdfDocumentResult(text).ToBytes());
        Assert.Equal(pages, pdf.NumberOfPages);
    }

    [Fact]
    public void MissingUnicodeGlyphsFailTheSaveInsteadOfPublishingReplacementText() {
        var conversion = PdfPlainTextConverter.ToPdfDocumentResult("漢字");
        using var output = new System.IO.MemoryStream();
        var result = conversion.SaveResult(output);
        Assert.False(result.Succeeded);
        Assert.True(result.Report.HasLoss);
    }

    [Fact]
    public void LiteralMarkupAndWhitespaceRenderWithoutInterpretation() {
        byte[] bytes = PdfPlainTextConverter.ToPdfDocumentResult("  # heading\r\n\r\n<b>literal</b>\tEND\fform feed").ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        string text = pdf.GetPage(1).Text;
        Assert.Contains("# heading", text);
        Assert.Contains("<b>literal</b>", text);
        Assert.Contains("END", text);
        var letters = pdf.GetPage(1).Letters;
        double headingX = letters.First(letter => letter.Value == "#").StartBaseLine.X;
        double literalX = letters.First(letter => letter.Value == "<").StartBaseLine.X;
        Assert.InRange(headingX - literalX, 11.5, 12.5); // Two Courier spaces at 10 points.
        Assert.True(letters.First(letter => letter.Value == "#").StartBaseLine.Y >
            letters.First(letter => letter.Value == "<").StartBaseLine.Y + 20);
    }

    [Theory]
    [InlineData("utf-8")]
    [InlineData("utf-16")]
    [InlineData("utf-32")]
    public void UnicodeBomSelectsEncoding(string name) {
        Encoding encoding = Encoding.GetEncoding(name);
        byte[] input = encoding.GetPreamble().Concat(encoding.GetBytes("BOM text")).ToArray();
        using var pdf = PdfPigDocument.Open(PdfPlainTextConverter.ToPdfDocumentResult(input).ToBytes());
        Assert.Contains("BOM text", pdf.GetPage(1).Text);
    }

    [Fact]
    public void InvalidDecodingAndExpandedTextCannotSilentlyLoseContent() {
        Assert.Throws<DecoderFallbackException>(() => PdfPlainTextConverter.ToPdfDocumentResult(new byte[] { 0xff }));
        Assert.Throws<System.IO.InvalidDataException>(() => PdfPlainTextConverter.ToPdfDocumentResult("\t", new PdfPlainTextOptions { MaximumCharacters = 2 }));
        byte[] utf16 = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("A")).ToArray();
        Assert.Throws<System.IO.InvalidDataException>(() => PdfPlainTextConverter.ToPdfDocumentResult(utf16, new PdfPlainTextOptions { EncodingName = "utf-8" }));
    }

    [Fact]
    public void GeneratedPageLimitIsEnforcedByTheRealWriter() {
        var conversion = PdfPlainTextConverter.ToPdfDocumentResult(new string('x', 4000), new PdfPlainTextOptions {
            MaximumPages = 1,
            PdfOptions = new PdfOptions { PageWidth = 120, PageHeight = 120, MarginLeft = 12, MarginRight = 12, MarginTop = 12, MarginBottom = 12 }
        });
        Assert.Throws<System.IO.InvalidDataException>(() => conversion.ToBytes());
    }
}

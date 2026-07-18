using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageExtractorTests {
    [Fact]
    public void ExtractPages_ReadsFromCurrentStreamPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        byte[] extracted = PdfPageExtractor.ExtractPages(stream, 2);

        var read = PdfReadDocument.Open(extracted);
        string text = NormalizeExtractedText(read.ExtractText());
        Assert.Single(read.Pages);
        Assert.Contains("Secondpagemarker", text);
        Assert.DoesNotContain("Firstpagemarker", text);
    }

    [Fact]
    public void ExtractPageRange_ReadsFromCurrentStreamPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        byte[] extracted = PdfPageExtractor.ExtractPageRange(stream, 1, 2);

        string text = NormalizeExtractedText(PdfReadDocument.Open(extracted).ExtractText());
        Assert.Contains("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.DoesNotContain("Thirdpagemarker", text);
    }

    [Fact]
    public void ExtractPageRanges_ReadsFromCurrentStreamPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        byte[] extracted = PdfPageExtractor.ExtractPageRanges(stream, PdfPageRange.From(3, 3), PdfPageRange.From(1, 1));

        string text = NormalizeExtractedText(PdfReadDocument.Open(extracted).ExtractText());
        AssertContainsInOrder(text, "Thirdpagemarker", "Firstpagemarker");
        Assert.DoesNotContain("Secondpagemarker", text);
    }

    [Fact]
    public void ExtractPages_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageExtractor.ExtractPages(source, output, 3);

        byte[] extracted = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        string text = NormalizeExtractedText(PdfReadDocument.Open(extracted).ExtractText());
        Assert.Contains("Thirdpagemarker", text);
        Assert.DoesNotContain("Firstpagemarker", text);
    }

    [Fact]
    public void ExtractPageRanges_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageExtractor.ExtractPageRanges(source, output, PdfPageRange.From(2, 2), PdfPageRange.From(1, 1));

        byte[] extracted = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        string text = NormalizeExtractedText(PdfReadDocument.Open(extracted).ExtractText());
        AssertContainsInOrder(text, "Secondpagemarker", "Firstpagemarker");
        Assert.DoesNotContain("Thirdpagemarker", text);
    }

    [Fact]
    public void ExtractPages_WritesStreamInputToOutputStream() {
        byte[] source = BuildThreePagePdf();
        byte[] inputPrefix = Encoding.ASCII.GetBytes("input-prefix");
        using var input = new MemoryStream(inputPrefix.Concat(source).ToArray());
        input.Position = inputPrefix.Length;
        using var output = new MemoryStream();

        PdfPageExtractor.ExtractPages(input, output, 2);

        string text = NormalizeExtractedText(PdfReadDocument.Open(output.ToArray()).ExtractText());
        Assert.Contains("Secondpagemarker", text);
        Assert.DoesNotContain("Firstpagemarker", text);
    }

    [Fact]
    public void ExtractPageRange_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageExtractor.ExtractPageRange(source, output, 1, 2);

        byte[] extracted = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        string text = NormalizeExtractedText(PdfReadDocument.Open(extracted).ExtractText());
        Assert.Contains("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.DoesNotContain("Thirdpagemarker", text);
    }

    [Fact]
    public void ExtractPageRange_WritesStreamInputToOutputStream() {
        byte[] source = BuildThreePagePdf();
        byte[] inputPrefix = Encoding.ASCII.GetBytes("input-prefix");
        using var input = new MemoryStream(inputPrefix.Concat(source).ToArray());
        input.Position = inputPrefix.Length;
        using var output = new MemoryStream();

        PdfPageExtractor.ExtractPageRange(input, output, 2, 3);

        string text = NormalizeExtractedText(PdfReadDocument.Open(output.ToArray()).ExtractText());
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }
}

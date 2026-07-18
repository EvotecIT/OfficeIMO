using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageEditorTests {

    [Fact]
    public void DeletePages_RemovesSelectedPagesAndKeepsOrder() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.DeletePages(source, 2);

        using var pdf = PdfPigDocument.Open(new MemoryStream(edited));
        Assert.Equal(2, pdf.NumberOfPages);

        var read = PdfReadDocument.Open(edited);
        string text = NormalizeExtractedText(read.ExtractText());
        Assert.Contains("Firstpagemarker", text);
        Assert.DoesNotContain("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
        Assert.True(text.IndexOf("Firstpagemarker", StringComparison.Ordinal) < text.IndexOf("Thirdpagemarker", StringComparison.Ordinal));

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Editor sample", info.Metadata.Title);
        Assert.Equal(2, info.PageCount);
        Assert.Equal(595, info.Pages[0].Width);
        Assert.Equal(842, info.Pages[0].Height);
        Assert.Equal(300, info.Pages[1].Width);
        Assert.Equal(500, info.Pages[1].Height);
    }

    [Fact]
    public void DeletePages_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-delete-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "deleted.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            PdfPageEditor.DeletePages(inputPath, outputPath, 1, 3);

            Assert.True(File.Exists(outputPath));
            PdfDocumentInfo info = PdfInspector.Inspect(outputPath);
            Assert.Equal(1, info.PageCount);
            string text = NormalizeExtractedText(PdfReadDocument.Open(outputPath).ExtractText());
            Assert.Contains("Secondpagemarker", text);
            Assert.DoesNotContain("Firstpagemarker", text);
            Assert.DoesNotContain("Thirdpagemarker", text);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void DeletePages_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.DeletePages(stream, 1, 3);

        var read = PdfReadDocument.Open(edited);
        string text = NormalizeExtractedText(read.ExtractText());
        Assert.Single(read.Pages);
        Assert.Contains("Secondpagemarker", text);
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.DoesNotContain("Thirdpagemarker", text);
    }

    [Fact]
    public void DeletePages_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.DeletePages(source, output, 1, 3);

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        string text = NormalizeExtractedText(PdfReadDocument.Open(edited).ExtractText());
        Assert.Contains("Secondpagemarker", text);
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.DoesNotContain("Thirdpagemarker", text);
    }

    [Fact]
    public void DeletePages_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.DeletePages(input, output, 2);

        string text = NormalizeExtractedText(PdfReadDocument.Open(output.ToArray()).ExtractText());
        Assert.Contains("Firstpagemarker", text);
        Assert.DoesNotContain("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }

    [Fact]
    public void DeletePageRange_RemovesInclusiveRange() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.DeletePageRange(source, 1, 2);

        using var pdf = PdfPigDocument.Open(new MemoryStream(edited));
        Assert.Equal(1, pdf.NumberOfPages);

        var read = PdfReadDocument.Open(edited);
        Assert.Single(read.Pages);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Editor sample", info.Metadata.Title);
        Assert.Equal(300, info.Pages[0].Width);
        Assert.Equal(500, info.Pages[0].Height);
    }

    [Fact]
    public void DeletePageRange_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-delete-range-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "deleted-range.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            PdfPageEditor.DeletePageRange(inputPath, outputPath, 2, 3);

            Assert.True(File.Exists(outputPath));
            var read = PdfReadDocument.Open(outputPath);
            Assert.Single(read.Pages);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void DeletePageRange_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.DeletePageRange(stream, 2, 3);

        var read = PdfReadDocument.Open(edited);
        Assert.Single(read.Pages);
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
    }

    [Fact]
    public void DeletePageRange_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.DeletePageRange(source, output, 1, 2);

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        var read = PdfReadDocument.Open(edited);
        Assert.Single(read.Pages);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
    }

    [Fact]
    public void DeletePageRange_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.DeletePageRange(input, output, 1, 1);

        var read = PdfReadDocument.Open(output.ToArray());
        Assert.Equal(2, read.Pages.Count);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void DeletePageRanges_RemovesParsedRangesAndAllowsOverlap() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.DeletePageRanges(source, PdfPageRange.ParseMany("1-2,2"));

        var read = PdfReadDocument.Open(edited);
        Assert.Single(read.Pages);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Editor sample", info.Metadata.Title);
        Assert.Equal(1, info.PageCount);
        Assert.Equal(300, info.Pages[0].Width);
        Assert.Equal(500, info.Pages[0].Height);
    }

    [Fact]
    public void DeletePageRanges_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.DeletePageRanges(stream, PdfPageRange.From(1, 1), PdfPageRange.From(3, 3));

        var read = PdfReadDocument.Open(edited);
        Assert.Single(read.Pages);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
    }

    [Fact]
    public void DeletePageRanges_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.DeletePageRanges(source, output, PdfPageRange.From(1, 1), PdfPageRange.From(3, 3));

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        var read = PdfReadDocument.Open(edited);
        Assert.Single(read.Pages);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
    }

    [Fact]
    public void DeletePageRanges_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-delete-ranges-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "deleted-ranges.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            PdfPageEditor.DeletePageRanges(inputPath, outputPath, PdfPageRange.From(1, 1), PdfPageRange.From(3, 3));

            Assert.True(File.Exists(outputPath));
            var read = PdfReadDocument.Open(outputPath);
            Assert.Single(read.Pages);
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

}

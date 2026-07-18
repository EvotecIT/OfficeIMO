using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageEditorTests {

    [Fact]
    public void DuplicatePages_ClonesSelectionsImmediatelyAfterSourcePages() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.DuplicatePages(source, 2, 2, 1);

        using var pdf = PdfPigDocument.Open(new MemoryStream(edited));
        Assert.Equal(6, pdf.NumberOfPages);

        var read = PdfReadDocument.Open(edited);
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[4].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[5].ExtractText()));

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Editor sample", info.Metadata.Title);
        Assert.Equal(6, info.PageCount);
        Assert.Equal(595, info.Pages[0].Width);
        Assert.Equal(842, info.Pages[0].Height);
        Assert.Equal(595, info.Pages[1].Width);
        Assert.Equal(842, info.Pages[1].Height);
        Assert.Equal(612, info.Pages[2].Width);
        Assert.Equal(792, info.Pages[2].Height);
        Assert.Equal(612, info.Pages[3].Width);
        Assert.Equal(792, info.Pages[3].Height);
        Assert.Equal(612, info.Pages[4].Width);
        Assert.Equal(792, info.Pages[4].Height);
        Assert.Equal(300, info.Pages[5].Width);
        Assert.Equal(500, info.Pages[5].Height);
    }

    [Fact]
    public void DuplicatePageRange_ClonesInclusiveRangeImmediatelyAfterSourcePages() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.DuplicatePageRange(source, 1, 2);

        using var pdf = PdfPigDocument.Open(new MemoryStream(edited));
        Assert.Equal(5, pdf.NumberOfPages);

        var read = PdfReadDocument.Open(edited);
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[4].ExtractText()));

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Editor sample", info.Metadata.Title);
        Assert.Equal(5, info.PageCount);
        Assert.Equal(595, info.Pages[0].Width);
        Assert.Equal(842, info.Pages[0].Height);
        Assert.Equal(595, info.Pages[1].Width);
        Assert.Equal(842, info.Pages[1].Height);
        Assert.Equal(612, info.Pages[2].Width);
        Assert.Equal(792, info.Pages[2].Height);
        Assert.Equal(612, info.Pages[3].Width);
        Assert.Equal(792, info.Pages[3].Height);
        Assert.Equal(300, info.Pages[4].Width);
        Assert.Equal(500, info.Pages[4].Height);
    }

    [Fact]
    public void DuplicatePageRanges_ClonesParsedRangesAndKeepsRepeatedOverlapCopies() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.DuplicatePageRanges(source, PdfPageRange.ParseMany("1-2,2"));

        var read = PdfReadDocument.Open(edited);
        Assert.Equal(6, read.Pages.Count);
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[4].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[5].ExtractText()));

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Editor sample", info.Metadata.Title);
        Assert.Equal(6, info.PageCount);
        Assert.Equal(595, info.Pages[0].Width);
        Assert.Equal(842, info.Pages[0].Height);
        Assert.Equal(595, info.Pages[1].Width);
        Assert.Equal(842, info.Pages[1].Height);
        Assert.Equal(612, info.Pages[2].Width);
        Assert.Equal(792, info.Pages[2].Height);
        Assert.Equal(612, info.Pages[3].Width);
        Assert.Equal(792, info.Pages[3].Height);
        Assert.Equal(612, info.Pages[4].Width);
        Assert.Equal(792, info.Pages[4].Height);
        Assert.Equal(300, info.Pages[5].Width);
        Assert.Equal(500, info.Pages[5].Height);
    }

    [Fact]
    public void DuplicatePages_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-duplicate-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "duplicated.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            PdfPageEditor.DuplicatePages(inputPath, outputPath, 3);

            Assert.True(File.Exists(outputPath));
            var read = PdfReadDocument.Open(outputPath);
            Assert.Equal(4, read.Pages.Count);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void DuplicatePageRange_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-duplicate-range-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "duplicated-range.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            PdfPageEditor.DuplicatePageRange(inputPath, outputPath, 2, 3);

            Assert.True(File.Exists(outputPath));
            var read = PdfReadDocument.Open(outputPath);
            Assert.Equal(5, read.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[4].ExtractText()));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void DuplicatePages_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.DuplicatePages(stream, 1);

        var read = PdfReadDocument.Open(edited);
        Assert.Equal(4, read.Pages.Count);
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void DuplicatePageRange_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.DuplicatePageRange(stream, 2, 3);

        var read = PdfReadDocument.Open(edited);
        Assert.Equal(5, read.Pages.Count);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[4].ExtractText()));
    }

    [Fact]
    public void DuplicatePageRanges_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.DuplicatePageRanges(stream, PdfPageRange.ParseMany("2-3,3"));

        var read = PdfReadDocument.Open(edited);
        Assert.Equal(6, read.Pages.Count);
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[4].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[5].ExtractText()));
    }

    [Fact]
    public void DuplicatePages_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.DuplicatePages(source, output, 2);

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        var read = PdfReadDocument.Open(edited);
        Assert.Equal(4, read.Pages.Count);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void DuplicatePageRange_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.DuplicatePageRange(source, output, 1, 2);

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        var read = PdfReadDocument.Open(edited);
        Assert.Equal(5, read.Pages.Count);
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
    }

    [Fact]
    public void DuplicatePageRanges_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.DuplicatePageRanges(source, output, PdfPageRange.From(1, 1), PdfPageRange.From(3, 3));

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        var read = PdfReadDocument.Open(edited);
        Assert.Equal(5, read.Pages.Count);
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[4].ExtractText()));
    }

    [Fact]
    public void DuplicatePages_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.DuplicatePages(input, output, 3);

        var read = PdfReadDocument.Open(output.ToArray());
        Assert.Equal(4, read.Pages.Count);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
    }

    [Fact]
    public void DuplicatePageRange_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.DuplicatePageRange(input, output, 2, 3);

        var read = PdfReadDocument.Open(output.ToArray());
        Assert.Equal(5, read.Pages.Count);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[4].ExtractText()));
    }

}

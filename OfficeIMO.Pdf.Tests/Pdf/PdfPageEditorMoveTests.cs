using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageEditorTests {

    [Fact]
    public void MovePages_MovesSelectionsBeforeTargetPageAndPreservesGeometry() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.MovePages(source, 1, 3);

        using var pdf = PdfPigDocument.Open(new MemoryStream(edited));
        Assert.Equal(3, pdf.NumberOfPages);

        var read = PdfReadDocument.Open(edited);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Editor sample", info.Metadata.Title);
        Assert.Equal(3, info.PageCount);
        Assert.Equal(300, info.Pages[0].Width);
        Assert.Equal(500, info.Pages[0].Height);
        Assert.Equal(595, info.Pages[1].Width);
        Assert.Equal(842, info.Pages[1].Height);
        Assert.Equal(612, info.Pages[2].Width);
        Assert.Equal(792, info.Pages[2].Height);
    }

    [Fact]
    public void MovePages_CanMoveMultiplePagesToEndInOriginalRelativeOrder() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.MovePages(source, 4, 1, 2);

        var read = PdfReadDocument.Open(edited);
        Assert.Equal(3, read.Pages.Count);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePageRange_MovesInclusiveRangeBeforeTargetPageAndPreservesGeometry() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.MovePageRange(source, 1, 2, 3);

        using var pdf = PdfPigDocument.Open(new MemoryStream(edited));
        Assert.Equal(3, pdf.NumberOfPages);

        var read = PdfReadDocument.Open(edited);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Editor sample", info.Metadata.Title);
        Assert.Equal(3, info.PageCount);
        Assert.Equal(612, info.Pages[0].Width);
        Assert.Equal(792, info.Pages[0].Height);
        Assert.Equal(300, info.Pages[1].Width);
        Assert.Equal(500, info.Pages[1].Height);
        Assert.Equal(595, info.Pages[2].Width);
        Assert.Equal(842, info.Pages[2].Height);
    }

    [Fact]
    public void MovePageRanges_MovesParsedRangesAndDeduplicatesOverlap() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.MovePageRanges(source, 4, PdfPageRange.ParseMany("1-2,2"));

        var read = PdfReadDocument.Open(edited);
        Assert.Equal(3, read.Pages.Count);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Editor sample", info.Metadata.Title);
        Assert.Equal(3, info.PageCount);
        Assert.Equal(300, info.Pages[0].Width);
        Assert.Equal(500, info.Pages[0].Height);
        Assert.Equal(595, info.Pages[1].Width);
        Assert.Equal(842, info.Pages[1].Height);
        Assert.Equal(612, info.Pages[2].Width);
        Assert.Equal(792, info.Pages[2].Height);
    }

    [Fact]
    public void MovePages_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-move-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "moved.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            PdfPageEditor.MovePages(inputPath, outputPath, 1, 2);

            Assert.True(File.Exists(outputPath));
            var read = PdfReadDocument.Open(outputPath);
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void MovePageRange_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-move-range-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "moved-range.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            PdfPageEditor.MovePageRange(inputPath, outputPath, 4, 1, 2);

            Assert.True(File.Exists(outputPath));
            var read = PdfReadDocument.Open(outputPath);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void MovePages_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.MovePages(stream, 1, 2);

        var read = PdfReadDocument.Open(edited);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePageRange_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.MovePageRange(stream, 1, 2, 3);

        var read = PdfReadDocument.Open(edited);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePageRanges_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.MovePageRanges(stream, 1, PdfPageRange.ParseMany("2-3,3"));

        var read = PdfReadDocument.Open(edited);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePages_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.MovePages(source, output, 4, 1);

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        var read = PdfReadDocument.Open(edited);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePageRange_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.MovePageRange(source, output, 4, 1, 2);

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        var read = PdfReadDocument.Open(edited);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePageRanges_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.MovePageRanges(source, output, 4, PdfPageRange.From(1, 1), PdfPageRange.From(2, 2));

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        var read = PdfReadDocument.Open(edited);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePages_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.MovePages(input, output, 4, 2);

        var read = PdfReadDocument.Open(output.ToArray());
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePageRange_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.MovePageRange(input, output, 1, 2, 3);

        var read = PdfReadDocument.Open(output.ToArray());
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

}

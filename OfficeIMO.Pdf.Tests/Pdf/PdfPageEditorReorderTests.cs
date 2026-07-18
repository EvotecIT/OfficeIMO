using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageEditorTests {

    [Fact]
    public void ReorderPages_ReordersAllPagesAndPreservesPageGeometry() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.ReorderPages(source, 3, 1, 2);

        using var pdf = PdfPigDocument.Open(new MemoryStream(edited));
        Assert.Equal(3, pdf.NumberOfPages);

        string text = NormalizeExtractedText(PdfReadDocument.Open(edited).ExtractText());
        Assert.Contains("Thirdpagemarker", text);
        Assert.Contains("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.True(text.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < text.IndexOf("Firstpagemarker", StringComparison.Ordinal));
        Assert.True(text.IndexOf("Firstpagemarker", StringComparison.Ordinal) < text.IndexOf("Secondpagemarker", StringComparison.Ordinal));

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
    public void ReorderPageRanges_ReordersExpandedRangesAndPreservesPageGeometry() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.ReorderPageRanges(source, new PdfPageRange(3, 3), new PdfPageRange(1, 2));

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
    public void ReorderPageRanges_AcceptsParsedRangeListForWrapperGrammar() {
        byte[] source = BuildThreePagePdf();
        PdfPageRange[] ranges = PdfPageRange.ParseMany("3,1-2");

        byte[] edited = PdfPageEditor.ReorderPageRanges(source, ranges);

        var read = PdfReadDocument.Open(edited);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void ReorderPages_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-reorder-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "reordered.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            PdfPageEditor.ReorderPages(inputPath, outputPath, 2, 3, 1);

            Assert.True(File.Exists(outputPath));
            string text = NormalizeExtractedText(PdfReadDocument.Open(outputPath).ExtractText());
            Assert.True(text.IndexOf("Secondpagemarker", StringComparison.Ordinal) < text.IndexOf("Thirdpagemarker", StringComparison.Ordinal));
            Assert.True(text.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < text.IndexOf("Firstpagemarker", StringComparison.Ordinal));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }


    [Fact]
    public void ReorderPages_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.ReorderPages(stream, 2, 3, 1);

        string text = NormalizeExtractedText(PdfReadDocument.Open(edited).ExtractText());
        Assert.True(text.IndexOf("Secondpagemarker", StringComparison.Ordinal) < text.IndexOf("Thirdpagemarker", StringComparison.Ordinal));
        Assert.True(text.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < text.IndexOf("Firstpagemarker", StringComparison.Ordinal));
    }

    [Fact]
    public void ReorderPages_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.ReorderPages(source, output, 3, 1, 2);

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        string text = NormalizeExtractedText(PdfReadDocument.Open(edited).ExtractText());
        Assert.True(text.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < text.IndexOf("Firstpagemarker", StringComparison.Ordinal));
        Assert.True(text.IndexOf("Firstpagemarker", StringComparison.Ordinal) < text.IndexOf("Secondpagemarker", StringComparison.Ordinal));
    }

    [Fact]
    public void ReorderPages_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.ReorderPages(input, output, 2, 3, 1);

        string text = NormalizeExtractedText(PdfReadDocument.Open(output.ToArray()).ExtractText());
        Assert.True(text.IndexOf("Secondpagemarker", StringComparison.Ordinal) < text.IndexOf("Thirdpagemarker", StringComparison.Ordinal));
        Assert.True(text.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < text.IndexOf("Firstpagemarker", StringComparison.Ordinal));
    }

}

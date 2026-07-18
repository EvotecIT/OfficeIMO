using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageEditorTests {

    [Fact]
    public void RotatePages_RotatesSelectedPages() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.RotatePages(source, 90, 2);

        using var pdf = PdfPigDocument.Open(new MemoryStream(edited));
        Assert.Equal(3, pdf.NumberOfPages);

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(0, info.Pages[0].RotationDegrees);
        Assert.Equal(90, info.Pages[1].RotationDegrees);
        Assert.Equal(0, info.Pages[2].RotationDegrees);

        string text = NormalizeExtractedText(PdfReadDocument.Open(edited).ExtractText());
        Assert.Contains("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }

    [Fact]
    public void RotatePages_RotatesAllPagesWhenNoSelectionIsProvidedAndNormalizesDegrees() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.RotatePages(source, -90);

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(270, info.Pages[0].RotationDegrees);
        Assert.Equal(270, info.Pages[1].RotationDegrees);
        Assert.Equal(270, info.Pages[2].RotationDegrees);
    }

    [Fact]
    public void RotatePageRange_RotatesInclusiveRange() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.RotatePageRange(source, 180, 2, 3);

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(0, info.Pages[0].RotationDegrees);
        Assert.Equal(180, info.Pages[1].RotationDegrees);
        Assert.Equal(180, info.Pages[2].RotationDegrees);

        string text = NormalizeExtractedText(PdfReadDocument.Open(edited).ExtractText());
        Assert.Contains("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }

    [Fact]
    public void RotatePageRanges_RotatesParsedRangesAndDeduplicatesOverlap() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.RotatePageRanges(source, 90, PdfPageRange.ParseMany("1-2,2"));

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(90, info.Pages[0].RotationDegrees);
        Assert.Equal(90, info.Pages[1].RotationDegrees);
        Assert.Equal(0, info.Pages[2].RotationDegrees);

        string text = NormalizeExtractedText(PdfReadDocument.Open(edited).ExtractText());
        Assert.Contains("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }

    [Fact]
    public void RotatePages_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-rotate-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "rotated.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            PdfPageEditor.RotatePages(inputPath, outputPath, 180, 1);

            Assert.True(File.Exists(outputPath));
            PdfDocumentInfo info = PdfInspector.Inspect(outputPath);
            Assert.Equal(180, info.Pages[0].RotationDegrees);
            Assert.Equal(0, info.Pages[1].RotationDegrees);
            Assert.Equal(0, info.Pages[2].RotationDegrees);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void RotatePageRange_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-rotate-range-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "rotated-range.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            PdfPageEditor.RotatePageRange(inputPath, outputPath, 270, 1, 2);

            Assert.True(File.Exists(outputPath));
            PdfDocumentInfo info = PdfInspector.Inspect(outputPath);
            Assert.Equal(270, info.Pages[0].RotationDegrees);
            Assert.Equal(270, info.Pages[1].RotationDegrees);
            Assert.Equal(0, info.Pages[2].RotationDegrees);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }


    [Fact]
    public void RotatePages_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.RotatePages(stream, 180, 3);

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(0, info.Pages[0].RotationDegrees);
        Assert.Equal(0, info.Pages[1].RotationDegrees);
        Assert.Equal(180, info.Pages[2].RotationDegrees);
    }

    [Fact]
    public void RotatePageRange_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.RotatePageRange(stream, 90, 1, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(90, info.Pages[0].RotationDegrees);
        Assert.Equal(90, info.Pages[1].RotationDegrees);
        Assert.Equal(0, info.Pages[2].RotationDegrees);
    }

    [Fact]
    public void RotatePageRanges_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.RotatePageRanges(stream, 180, PdfPageRange.ParseMany("2-3,3"));

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(0, info.Pages[0].RotationDegrees);
        Assert.Equal(180, info.Pages[1].RotationDegrees);
        Assert.Equal(180, info.Pages[2].RotationDegrees);
    }

    [Fact]
    public void RotatePages_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.RotatePages(source, output, 180, 1);

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(180, info.Pages[0].RotationDegrees);
        Assert.Equal(0, info.Pages[1].RotationDegrees);
        Assert.Equal(0, info.Pages[2].RotationDegrees);
    }

    [Fact]
    public void RotatePageRange_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.RotatePageRange(source, output, 180, 2, 3);

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(0, info.Pages[0].RotationDegrees);
        Assert.Equal(180, info.Pages[1].RotationDegrees);
        Assert.Equal(180, info.Pages[2].RotationDegrees);
    }

    [Fact]
    public void RotatePageRanges_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageEditor.RotatePageRanges(source, output, 270, PdfPageRange.From(1, 1), PdfPageRange.From(3, 3));

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(270, info.Pages[0].RotationDegrees);
        Assert.Equal(0, info.Pages[1].RotationDegrees);
        Assert.Equal(270, info.Pages[2].RotationDegrees);
    }

    [Fact]
    public void RotatePages_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.RotatePages(input, output, 270, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(output.ToArray());
        Assert.Equal(0, info.Pages[0].RotationDegrees);
        Assert.Equal(270, info.Pages[1].RotationDegrees);
        Assert.Equal(0, info.Pages[2].RotationDegrees);
    }

    [Fact]
    public void RotatePageRange_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.RotatePageRange(input, output, 270, 1, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(output.ToArray());
        Assert.Equal(270, info.Pages[0].RotationDegrees);
        Assert.Equal(270, info.Pages[1].RotationDegrees);
        Assert.Equal(0, info.Pages[2].RotationDegrees);
    }

}

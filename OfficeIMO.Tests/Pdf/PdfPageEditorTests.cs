using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageEditorTests {
    [Fact]
    public void DeletePages_RemovesSelectedPagesAndKeepsOrder() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.DeletePages(source, 2);

        using var pdf = PdfDocument.Open(new MemoryStream(edited));
        Assert.Equal(2, pdf.NumberOfPages);

        var read = PdfReadDocument.Load(edited);
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
            string text = NormalizeExtractedText(PdfReadDocument.Load(outputPath).ExtractText());
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

        var read = PdfReadDocument.Load(edited);
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

        string text = NormalizeExtractedText(PdfReadDocument.Load(edited).ExtractText());
        Assert.Contains("Secondpagemarker", text);
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.DoesNotContain("Thirdpagemarker", text);
    }

    [Fact]
    public void DeletePages_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.DeletePages(input, output, 2);

        string text = NormalizeExtractedText(PdfReadDocument.Load(output.ToArray()).ExtractText());
        Assert.Contains("Firstpagemarker", text);
        Assert.DoesNotContain("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }

    [Fact]
    public void DeletePageRange_RemovesInclusiveRange() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.DeletePageRange(source, 1, 2);

        using var pdf = PdfDocument.Open(new MemoryStream(edited));
        Assert.Equal(1, pdf.NumberOfPages);

        var read = PdfReadDocument.Load(edited);
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
            var read = PdfReadDocument.Load(outputPath);
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

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
        Assert.Single(read.Pages);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
    }

    [Fact]
    public void DeletePageRange_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.DeletePageRange(input, output, 1, 1);

        var read = PdfReadDocument.Load(output.ToArray());
        Assert.Equal(2, read.Pages.Count);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void DeletePageRanges_RemovesParsedRangesAndAllowsOverlap() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.DeletePageRanges(source, PdfPageRange.ParseMany("1-2,2"));

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
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
            var read = PdfReadDocument.Load(outputPath);
            Assert.Single(read.Pages);
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void DuplicatePages_ClonesSelectionsImmediatelyAfterSourcePages() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.DuplicatePages(source, 2, 2, 1);

        using var pdf = PdfDocument.Open(new MemoryStream(edited));
        Assert.Equal(6, pdf.NumberOfPages);

        var read = PdfReadDocument.Load(edited);
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

        using var pdf = PdfDocument.Open(new MemoryStream(edited));
        Assert.Equal(5, pdf.NumberOfPages);

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
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
            var read = PdfReadDocument.Load(outputPath);
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
            var read = PdfReadDocument.Load(outputPath);
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

        var read = PdfReadDocument.Load(edited);
        Assert.Equal(4, read.Pages.Count);
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void DuplicatePageRange_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.DuplicatePageRange(stream, 2, 3);

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(output.ToArray());
        Assert.Equal(4, read.Pages.Count);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
    }

    [Fact]
    public void DuplicatePageRange_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.DuplicatePageRange(input, output, 2, 3);

        var read = PdfReadDocument.Load(output.ToArray());
        Assert.Equal(5, read.Pages.Count);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[3].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[4].ExtractText()));
    }

    [Fact]
    public void MovePages_MovesSelectionsBeforeTargetPageAndPreservesGeometry() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.MovePages(source, 1, 3);

        using var pdf = PdfDocument.Open(new MemoryStream(edited));
        Assert.Equal(3, pdf.NumberOfPages);

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
        Assert.Equal(3, read.Pages.Count);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePageRange_MovesInclusiveRangeBeforeTargetPageAndPreservesGeometry() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.MovePageRange(source, 1, 2, 3);

        using var pdf = PdfDocument.Open(new MemoryStream(edited));
        Assert.Equal(3, pdf.NumberOfPages);

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
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
            var read = PdfReadDocument.Load(outputPath);
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
            var read = PdfReadDocument.Load(outputPath);
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

        var read = PdfReadDocument.Load(edited);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePageRange_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.MovePageRange(stream, 1, 2, 3);

        var read = PdfReadDocument.Load(edited);
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePageRanges_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildThreePagePdf());

        byte[] edited = PdfPageEditor.MovePageRanges(stream, 1, PdfPageRange.ParseMany("2-3,3"));

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePages_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.MovePages(input, output, 4, 2);

        var read = PdfReadDocument.Load(output.ToArray());
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void MovePageRange_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.MovePageRange(input, output, 1, 2, 3);

        var read = PdfReadDocument.Load(output.ToArray());
        Assert.Equal("Secondpagemarker", NormalizeExtractedText(read.Pages[0].ExtractText()));
        Assert.Equal("Thirdpagemarker", NormalizeExtractedText(read.Pages[1].ExtractText()));
        Assert.Equal("Firstpagemarker", NormalizeExtractedText(read.Pages[2].ExtractText()));
    }

    [Fact]
    public void ReorderPages_ReordersAllPagesAndPreservesPageGeometry() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.ReorderPages(source, 3, 1, 2);

        using var pdf = PdfDocument.Open(new MemoryStream(edited));
        Assert.Equal(3, pdf.NumberOfPages);

        string text = NormalizeExtractedText(PdfReadDocument.Load(edited).ExtractText());
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

        var read = PdfReadDocument.Load(edited);
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

        var read = PdfReadDocument.Load(edited);
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
            string text = NormalizeExtractedText(PdfReadDocument.Load(outputPath).ExtractText());
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

        string text = NormalizeExtractedText(PdfReadDocument.Load(edited).ExtractText());
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

        string text = NormalizeExtractedText(PdfReadDocument.Load(edited).ExtractText());
        Assert.True(text.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < text.IndexOf("Firstpagemarker", StringComparison.Ordinal));
        Assert.True(text.IndexOf("Firstpagemarker", StringComparison.Ordinal) < text.IndexOf("Secondpagemarker", StringComparison.Ordinal));
    }

    [Fact]
    public void ReorderPages_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildThreePagePdf());
        using var output = new MemoryStream();

        PdfPageEditor.ReorderPages(input, output, 2, 3, 1);

        string text = NormalizeExtractedText(PdfReadDocument.Load(output.ToArray()).ExtractText());
        Assert.True(text.IndexOf("Secondpagemarker", StringComparison.Ordinal) < text.IndexOf("Thirdpagemarker", StringComparison.Ordinal));
        Assert.True(text.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < text.IndexOf("Firstpagemarker", StringComparison.Ordinal));
    }

    [Fact]
    public void RotatePages_RotatesSelectedPages() {
        byte[] source = BuildThreePagePdf();

        byte[] edited = PdfPageEditor.RotatePages(source, 90, 2);

        using var pdf = PdfDocument.Open(new MemoryStream(edited));
        Assert.Equal(3, pdf.NumberOfPages);

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(0, info.Pages[0].RotationDegrees);
        Assert.Equal(90, info.Pages[1].RotationDegrees);
        Assert.Equal(0, info.Pages[2].RotationDegrees);

        string text = NormalizeExtractedText(PdfReadDocument.Load(edited).ExtractText());
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

        string text = NormalizeExtractedText(PdfReadDocument.Load(edited).ExtractText());
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

        string text = NormalizeExtractedText(PdfReadDocument.Load(edited).ExtractText());
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
    public void PageEditorPathInputs_ReturnBytesForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-editor-path-bytes-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            byte[] deleted = PdfPageEditor.DeletePages(inputPath, 2);
            string deletedText = NormalizeExtractedText(PdfReadDocument.Load(deleted).ExtractText());
            Assert.Contains("Firstpagemarker", deletedText);
            Assert.DoesNotContain("Secondpagemarker", deletedText);
            Assert.Contains("Thirdpagemarker", deletedText);

            byte[] deletedRange = PdfPageEditor.DeletePageRange(inputPath, 1, 2);
            var deletedRangeRead = PdfReadDocument.Load(deletedRange);
            Assert.Single(deletedRangeRead.Pages);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(deletedRangeRead.Pages[0].ExtractText()));

            byte[] deletedModelRange = PdfPageEditor.DeletePageRange(inputPath, PdfPageRange.From(1, 2));
            var deletedModelRangeRead = PdfReadDocument.Load(deletedModelRange);
            Assert.Single(deletedModelRangeRead.Pages);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(deletedModelRangeRead.Pages[0].ExtractText()));

            byte[] deletedRanges = PdfPageEditor.DeletePageRanges(inputPath, PdfPageRange.From(1, 1), PdfPageRange.From(3, 3));
            var deletedRangesRead = PdfReadDocument.Load(deletedRanges);
            Assert.Single(deletedRangesRead.Pages);
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(deletedRangesRead.Pages[0].ExtractText()));

            byte[] duplicated = PdfPageEditor.DuplicatePages(inputPath, 3);
            var duplicatedRead = PdfReadDocument.Load(duplicated);
            Assert.Equal(4, duplicatedRead.Pages.Count);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRead.Pages[2].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRead.Pages[3].ExtractText()));

            byte[] duplicatedRange = PdfPageEditor.DuplicatePageRange(inputPath, 1, 2);
            var duplicatedRangeRead = PdfReadDocument.Load(duplicatedRange);
            Assert.Equal(5, duplicatedRangeRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[2].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[3].ExtractText()));

            byte[] duplicatedModelRange = PdfPageEditor.DuplicatePageRange(inputPath, PdfPageRange.From(1, 2));
            var duplicatedModelRangeRead = PdfReadDocument.Load(duplicatedModelRange);
            Assert.Equal(5, duplicatedModelRangeRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[2].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[3].ExtractText()));

            byte[] duplicatedRanges = PdfPageEditor.DuplicatePageRanges(inputPath, PdfPageRange.ParseMany("1,3"));
            var duplicatedRangesRead = PdfReadDocument.Load(duplicatedRanges);
            Assert.Equal(5, duplicatedRangesRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[2].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[3].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[4].ExtractText()));

            byte[] moved = PdfPageEditor.MovePages(inputPath, 1, 2);
            var movedRead = PdfReadDocument.Load(moved);
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRead.Pages[1].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRead.Pages[2].ExtractText()));

            byte[] movedRange = PdfPageEditor.MovePageRange(inputPath, 4, 1, 2);
            var movedRangeRead = PdfReadDocument.Load(movedRange);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRangeRead.Pages[2].ExtractText()));

            byte[] movedModelRange = PdfPageEditor.MovePageRange(inputPath, 4, PdfPageRange.From(1, 2));
            var movedModelRangeRead = PdfReadDocument.Load(movedModelRange);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[2].ExtractText()));

            byte[] movedRanges = PdfPageEditor.MovePageRanges(inputPath, 4, PdfPageRange.ParseMany("1-2,2"));
            var movedRangesRead = PdfReadDocument.Load(movedRanges);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRangesRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRangesRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRangesRead.Pages[2].ExtractText()));

            byte[] reordered = PdfPageEditor.ReorderPages(inputPath, 3, 1, 2);
            string reorderedText = NormalizeExtractedText(PdfReadDocument.Load(reordered).ExtractText());
            Assert.True(reorderedText.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < reorderedText.IndexOf("Firstpagemarker", StringComparison.Ordinal));
            Assert.True(reorderedText.IndexOf("Firstpagemarker", StringComparison.Ordinal) < reorderedText.IndexOf("Secondpagemarker", StringComparison.Ordinal));

            byte[] rotated = PdfPageEditor.RotatePages(inputPath, 90, 1);
            PdfDocumentInfo rotatedInfo = PdfInspector.Inspect(rotated);
            Assert.Equal(90, rotatedInfo.Pages[0].RotationDegrees);
            Assert.Equal(0, rotatedInfo.Pages[1].RotationDegrees);
            Assert.Equal(0, rotatedInfo.Pages[2].RotationDegrees);

            byte[] rotatedRange = PdfPageEditor.RotatePageRange(inputPath, 180, 2, 3);
            PdfDocumentInfo rotatedRangeInfo = PdfInspector.Inspect(rotatedRange);
            Assert.Equal(0, rotatedRangeInfo.Pages[0].RotationDegrees);
            Assert.Equal(180, rotatedRangeInfo.Pages[1].RotationDegrees);
            Assert.Equal(180, rotatedRangeInfo.Pages[2].RotationDegrees);

            byte[] rotatedModelRange = PdfPageEditor.RotatePageRange(inputPath, 180, PdfPageRange.From(2, 3));
            PdfDocumentInfo rotatedModelRangeInfo = PdfInspector.Inspect(rotatedModelRange);
            Assert.Equal(0, rotatedModelRangeInfo.Pages[0].RotationDegrees);
            Assert.Equal(180, rotatedModelRangeInfo.Pages[1].RotationDegrees);
            Assert.Equal(180, rotatedModelRangeInfo.Pages[2].RotationDegrees);

            byte[] rotatedRanges = PdfPageEditor.RotatePageRanges(inputPath, 270, PdfPageRange.ParseMany("1,3"));
            PdfDocumentInfo rotatedRangesInfo = PdfInspector.Inspect(rotatedRanges);
            Assert.Equal(270, rotatedRangesInfo.Pages[0].RotationDegrees);
            Assert.Equal(0, rotatedRangesInfo.Pages[1].RotationDegrees);
            Assert.Equal(270, rotatedRangesInfo.Pages[2].RotationDegrees);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void PageEditorPathInputs_WriteToOutputStreamsForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-editor-path-stream-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            using var deletedOutput = CreateOutputStream(out int deletedPrefixLength);
            PdfPageEditor.DeletePages(inputPath, deletedOutput, 2);
            string deletedText = NormalizeExtractedText(PdfReadDocument.Load(GetOutputPayload(deletedOutput, deletedPrefixLength)).ExtractText());
            Assert.Contains("Firstpagemarker", deletedText);
            Assert.DoesNotContain("Secondpagemarker", deletedText);
            Assert.Contains("Thirdpagemarker", deletedText);

            using var deletedRangeOutput = CreateOutputStream(out int deletedRangePrefixLength);
            PdfPageEditor.DeletePageRange(inputPath, deletedRangeOutput, 1, 2);
            var deletedRangeRead = PdfReadDocument.Load(GetOutputPayload(deletedRangeOutput, deletedRangePrefixLength));
            Assert.Single(deletedRangeRead.Pages);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(deletedRangeRead.Pages[0].ExtractText()));

            using var deletedModelRangeOutput = CreateOutputStream(out int deletedModelRangePrefixLength);
            PdfPageEditor.DeletePageRange(inputPath, deletedModelRangeOutput, PdfPageRange.From(1, 2));
            var deletedModelRangeRead = PdfReadDocument.Load(GetOutputPayload(deletedModelRangeOutput, deletedModelRangePrefixLength));
            Assert.Single(deletedModelRangeRead.Pages);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(deletedModelRangeRead.Pages[0].ExtractText()));

            using var deletedRangesOutput = CreateOutputStream(out int deletedRangesPrefixLength);
            PdfPageEditor.DeletePageRanges(inputPath, deletedRangesOutput, PdfPageRange.From(1, 1), PdfPageRange.From(3, 3));
            var deletedRangesRead = PdfReadDocument.Load(GetOutputPayload(deletedRangesOutput, deletedRangesPrefixLength));
            Assert.Single(deletedRangesRead.Pages);
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(deletedRangesRead.Pages[0].ExtractText()));

            using var duplicatedOutput = CreateOutputStream(out int duplicatedPrefixLength);
            PdfPageEditor.DuplicatePages(inputPath, duplicatedOutput, 3);
            var duplicatedRead = PdfReadDocument.Load(GetOutputPayload(duplicatedOutput, duplicatedPrefixLength));
            Assert.Equal(4, duplicatedRead.Pages.Count);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRead.Pages[2].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRead.Pages[3].ExtractText()));

            using var duplicatedRangeOutput = CreateOutputStream(out int duplicatedRangePrefixLength);
            PdfPageEditor.DuplicatePageRange(inputPath, duplicatedRangeOutput, 1, 2);
            var duplicatedRangeRead = PdfReadDocument.Load(GetOutputPayload(duplicatedRangeOutput, duplicatedRangePrefixLength));
            Assert.Equal(5, duplicatedRangeRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[2].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[3].ExtractText()));

            using var duplicatedModelRangeOutput = CreateOutputStream(out int duplicatedModelRangePrefixLength);
            PdfPageEditor.DuplicatePageRange(inputPath, duplicatedModelRangeOutput, PdfPageRange.From(1, 2));
            var duplicatedModelRangeRead = PdfReadDocument.Load(GetOutputPayload(duplicatedModelRangeOutput, duplicatedModelRangePrefixLength));
            Assert.Equal(5, duplicatedModelRangeRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[2].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[3].ExtractText()));

            using var duplicatedRangesOutput = CreateOutputStream(out int duplicatedRangesPrefixLength);
            PdfPageEditor.DuplicatePageRanges(inputPath, duplicatedRangesOutput, PdfPageRange.ParseMany("1,3"));
            var duplicatedRangesRead = PdfReadDocument.Load(GetOutputPayload(duplicatedRangesOutput, duplicatedRangesPrefixLength));
            Assert.Equal(5, duplicatedRangesRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[2].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[3].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[4].ExtractText()));

            using var movedOutput = CreateOutputStream(out int movedPrefixLength);
            PdfPageEditor.MovePages(inputPath, movedOutput, 1, 2);
            var movedRead = PdfReadDocument.Load(GetOutputPayload(movedOutput, movedPrefixLength));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRead.Pages[1].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRead.Pages[2].ExtractText()));

            using var movedRangeOutput = CreateOutputStream(out int movedRangePrefixLength);
            PdfPageEditor.MovePageRange(inputPath, movedRangeOutput, 4, 1, 2);
            var movedRangeRead = PdfReadDocument.Load(GetOutputPayload(movedRangeOutput, movedRangePrefixLength));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRangeRead.Pages[2].ExtractText()));

            using var movedModelRangeOutput = CreateOutputStream(out int movedModelRangePrefixLength);
            PdfPageEditor.MovePageRange(inputPath, movedModelRangeOutput, 4, PdfPageRange.From(1, 2));
            var movedModelRangeRead = PdfReadDocument.Load(GetOutputPayload(movedModelRangeOutput, movedModelRangePrefixLength));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[2].ExtractText()));

            using var movedRangesOutput = CreateOutputStream(out int movedRangesPrefixLength);
            PdfPageEditor.MovePageRanges(inputPath, movedRangesOutput, 4, PdfPageRange.ParseMany("1-2,2"));
            var movedRangesRead = PdfReadDocument.Load(GetOutputPayload(movedRangesOutput, movedRangesPrefixLength));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRangesRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRangesRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRangesRead.Pages[2].ExtractText()));

            using var reorderedOutput = CreateOutputStream(out int reorderedPrefixLength);
            PdfPageEditor.ReorderPages(inputPath, reorderedOutput, 3, 1, 2);
            string reorderedText = NormalizeExtractedText(PdfReadDocument.Load(GetOutputPayload(reorderedOutput, reorderedPrefixLength)).ExtractText());
            Assert.True(reorderedText.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < reorderedText.IndexOf("Firstpagemarker", StringComparison.Ordinal));
            Assert.True(reorderedText.IndexOf("Firstpagemarker", StringComparison.Ordinal) < reorderedText.IndexOf("Secondpagemarker", StringComparison.Ordinal));

            using var reorderedRangesOutput = CreateOutputStream(out int reorderedRangesPrefixLength);
            PdfPageEditor.ReorderPageRanges(inputPath, reorderedRangesOutput, PdfPageRange.From(3, 3), PdfPageRange.From(1, 2));
            var reorderedRangesRead = PdfReadDocument.Load(GetOutputPayload(reorderedRangesOutput, reorderedRangesPrefixLength));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(reorderedRangesRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(reorderedRangesRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(reorderedRangesRead.Pages[2].ExtractText()));

            using var rotatedOutput = CreateOutputStream(out int rotatedPrefixLength);
            PdfPageEditor.RotatePages(inputPath, rotatedOutput, 90, 1);
            PdfDocumentInfo rotatedInfo = PdfInspector.Inspect(GetOutputPayload(rotatedOutput, rotatedPrefixLength));
            Assert.Equal(90, rotatedInfo.Pages[0].RotationDegrees);
            Assert.Equal(0, rotatedInfo.Pages[1].RotationDegrees);
            Assert.Equal(0, rotatedInfo.Pages[2].RotationDegrees);

            using var rotatedRangeOutput = CreateOutputStream(out int rotatedRangePrefixLength);
            PdfPageEditor.RotatePageRange(inputPath, rotatedRangeOutput, 180, 2, 3);
            PdfDocumentInfo rotatedRangeInfo = PdfInspector.Inspect(GetOutputPayload(rotatedRangeOutput, rotatedRangePrefixLength));
            Assert.Equal(0, rotatedRangeInfo.Pages[0].RotationDegrees);
            Assert.Equal(180, rotatedRangeInfo.Pages[1].RotationDegrees);
            Assert.Equal(180, rotatedRangeInfo.Pages[2].RotationDegrees);

            using var rotatedModelRangeOutput = CreateOutputStream(out int rotatedModelRangePrefixLength);
            PdfPageEditor.RotatePageRange(inputPath, rotatedModelRangeOutput, 180, PdfPageRange.From(2, 3));
            PdfDocumentInfo rotatedModelRangeInfo = PdfInspector.Inspect(GetOutputPayload(rotatedModelRangeOutput, rotatedModelRangePrefixLength));
            Assert.Equal(0, rotatedModelRangeInfo.Pages[0].RotationDegrees);
            Assert.Equal(180, rotatedModelRangeInfo.Pages[1].RotationDegrees);
            Assert.Equal(180, rotatedModelRangeInfo.Pages[2].RotationDegrees);

            using var rotatedRangesOutput = CreateOutputStream(out int rotatedRangesPrefixLength);
            PdfPageEditor.RotatePageRanges(inputPath, rotatedRangesOutput, 270, PdfPageRange.ParseMany("1,3"));
            PdfDocumentInfo rotatedRangesInfo = PdfInspector.Inspect(GetOutputPayload(rotatedRangesOutput, rotatedRangesPrefixLength));
            Assert.Equal(270, rotatedRangesInfo.Pages[0].RotationDegrees);
            Assert.Equal(0, rotatedRangesInfo.Pages[1].RotationDegrees);
            Assert.Equal(270, rotatedRangesInfo.Pages[2].RotationDegrees);
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

    [Fact]
    public void PageEditor_RejectsInvalidSelections() {
        byte[] source = BuildThreePagePdf();

        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(source, Array.Empty<int>()));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(source, 1, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(source, 1, 2, 3));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePages(source, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePages(source, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePages((Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(new WriteOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePages(source, null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(source, new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePages(new MemoryStream(source), null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(new MemoryStream(source), new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePages("input.pdf", (Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages("missing.pdf", new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(" ", new MemoryStream(), 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(" ", "out.pdf", 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages("missing.pdf", " ", 1));

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePageRange(source, 3, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(source, 1, 3));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePageRange(source, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePageRange(source, 1, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRange((Stream)null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(new WriteOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRange(source, null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(source, new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRange(new MemoryStream(source), null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(new MemoryStream(source), new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRange("input.pdf", (Stream)null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange("missing.pdf", new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(" ", new MemoryStream(), 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(" ", "out.pdf", 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange("missing.pdf", " ", 1, 2));

        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRanges(source, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(source, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(source, PdfPageRange.From(1, 3)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePageRanges(source, PdfPageRange.From(1, 2), PdfPageRange.From(4, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRanges((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(new WriteOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRanges(source, null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(source, new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRanges(new MemoryStream(source), null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(new MemoryStream(source), new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(" ", "out.pdf", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges("missing.pdf", " ", PdfPageRange.From(1, 1)));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(source, Array.Empty<int>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePages(source, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePages(source, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePages((Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(new WriteOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePages(source, null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(source, new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePages(new MemoryStream(source), null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(new MemoryStream(source), new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePages("input.pdf", (Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages("missing.pdf", new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(" ", new MemoryStream(), 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(" ", "out.pdf", 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages("missing.pdf", " ", 1));

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePageRange(source, 3, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePageRange(source, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePageRange(source, 1, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRange((Stream)null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRange(new WriteOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRange(source, null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRange(source, new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRange(new MemoryStream(source), null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRange(new MemoryStream(source), new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRange(" ", "out.pdf", 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRange("missing.pdf", " ", 1, 2));

        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRanges(source, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges(source, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePageRanges(source, PdfPageRange.From(1, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRanges((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges(new WriteOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRanges(source, null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges(source, new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRanges(new MemoryStream(source), null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges(new MemoryStream(source), new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges(" ", "out.pdf", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges("missing.pdf", " ", PdfPageRange.From(1, 1)));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(source, 1, Array.Empty<int>()));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(source, 2, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(source, 1, 1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePages(source, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePages(source, 5, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePages(source, 1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePages(source, 1, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePages((Stream)null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(new WriteOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePages(source, null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(source, new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePages(new MemoryStream(source), null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(new MemoryStream(source), new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePages("input.pdf", (Stream)null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages("missing.pdf", new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(" ", new MemoryStream(), 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(" ", "out.pdf", 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages("missing.pdf", " ", 1, 2));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange(source, 2, 2, 3));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRange(source, 1, 3, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRange(source, 0, 1, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRange(source, 5, 1, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRange(source, 1, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRange(source, 1, 1, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRange((Stream)null!, 1, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange(new WriteOnlyStream(), 1, 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRange(source, null!, 4, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange(source, new ReadOnlyStream(), 4, 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRange(new MemoryStream(source), null!, 4, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange(new MemoryStream(source), new ReadOnlyStream(), 4, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange(" ", "out.pdf", 1, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange("missing.pdf", " ", 1, 1, 2));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(source, 2, PdfPageRange.From(2, 3)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRanges((byte[])null!, 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRanges(source, 4, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(source, 4, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRanges(source, 4, PdfPageRange.From(1, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRanges((Stream)null!, 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(new WriteOnlyStream(), 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRanges(source, null!, 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(source, new ReadOnlyStream(), 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRanges(new MemoryStream(source), null!, 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(new MemoryStream(source), new ReadOnlyStream(), 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(" ", "out.pdf", 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges("missing.pdf", " ", 4, PdfPageRange.From(1, 1)));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(source, Array.Empty<int>()));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(source, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(source, 1, 2, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.ReorderPages(source, 1, 2, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.ReorderPages(source, 1, 2, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPages((Stream)null!, 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(new WriteOnlyStream(), 1, 2, 3));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPages(source, null!, 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(source, new ReadOnlyStream(), 1, 2, 3));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPages(new MemoryStream(source), null!, 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(new MemoryStream(source), new ReadOnlyStream(), 1, 2, 3));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPages("input.pdf", (Stream)null!, 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages("missing.pdf", new ReadOnlyStream(), 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(" ", new MemoryStream(), 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(" ", "out.pdf", 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages("missing.pdf", " ", 1, 2, 3));

        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPageRanges(source, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(source, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(source, new PdfPageRange(3, 3), new PdfPageRange(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(source, new PdfPageRange(1, 2), new PdfPageRange(2, 3)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.ReorderPageRanges(source, new PdfPageRange(1, 2), new PdfPageRange(4, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPageRanges((Stream)null!, new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(new WriteOnlyStream(), new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPageRanges(source, null!, new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(source, new ReadOnlyStream(), new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPageRanges(new MemoryStream(source), null!, new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(new MemoryStream(source), new ReadOnlyStream(), new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(" ", "out.pdf", new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges("missing.pdf", " ", new PdfPageRange(1, 3)));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(source, 90, 1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePages(source, 45, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePages(source, 90, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePages(source, 90, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePages((Stream)null!, 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(new WriteOnlyStream(), 90, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePages(source, null!, 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(source, new ReadOnlyStream(), 90, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePages(new MemoryStream(source), null!, 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(new MemoryStream(source), new ReadOnlyStream(), 90, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePages("input.pdf", (Stream)null!, 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages("missing.pdf", new ReadOnlyStream(), 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(" ", new MemoryStream(), 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(" ", "out.pdf", 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages("missing.pdf", " ", 90, 1));

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRange(source, 90, 3, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRange(source, 90, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRange(source, 90, 1, 4));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRange(source, 45, 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRange((Stream)null!, 90, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRange(new WriteOnlyStream(), 90, 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRange(source, null!, 90, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRange(source, new ReadOnlyStream(), 90, 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRange(new MemoryStream(source), null!, 90, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRange(new MemoryStream(source), new ReadOnlyStream(), 90, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRange(" ", "out.pdf", 90, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRange("missing.pdf", " ", 90, 1, 2));

        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRanges((byte[])null!, 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRanges(source, 90, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges(source, 90, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRanges(source, 90, PdfPageRange.From(1, 4)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRanges(source, 45, PdfPageRange.From(1, 2)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRanges((Stream)null!, 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges(new WriteOnlyStream(), 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRanges(source, null!, 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges(source, new ReadOnlyStream(), 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRanges(new MemoryStream(source), null!, 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges(new MemoryStream(source), new ReadOnlyStream(), 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges(" ", "out.pdf", 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges("missing.pdf", " ", 90, PdfPageRange.From(1, 1)));
    }

    [Fact]
    public void PageEditorPathOutputs_RejectDirectoryTargets() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-page-editor-output-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputDirectory = Path.Combine(directory, "existing-output");

        try {
            Directory.CreateDirectory(outputDirectory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            var ex = Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(inputPath, outputDirectory, 90));
            Assert.Equal("outputPath", ex.ParamName);
            Assert.Contains("Output path refers to a directory; a file path is required.", ex.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static byte[] BuildThreePagePdf() {
        var doc = PdfDoc.Create()
            .Meta(title: "Editor sample", author: "OfficeIMO");

        doc.Compose(compose => {
            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("First page marker"))));
            });

            compose.Page(page => {
                page.Size(new PageSize(612, 792));
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Second page marker"))));
            });

            compose.Page(page => {
                page.Size(new PageSize(300, 500));
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Third page marker"))));
            });
        });

        return doc.ToBytes();
    }

    private static string NormalizeExtractedText(string text) {
        return text.Replace(" ", string.Empty);
    }

    private static MemoryStream CreatePrefixedStream(byte[] pdf) {
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        var stream = new MemoryStream();
        stream.Write(prefix, 0, prefix.Length);
        stream.Write(pdf, 0, pdf.Length);
        stream.Position = prefix.Length;
        return stream;
    }

    private static MemoryStream CreateOutputStream(out int prefixLength) {
        byte[] prefix = Encoding.ASCII.GetBytes("output-prefix");
        var stream = new MemoryStream();
        stream.Write(prefix, 0, prefix.Length);
        prefixLength = prefix.Length;
        return stream;
    }

    private static byte[] GetOutputPayload(MemoryStream output, int prefixLength) {
        byte[] bytes = output.ToArray();
        Assert.True(bytes.Length > prefixLength);
        Assert.Equal("output-prefix", Encoding.ASCII.GetString(bytes, 0, prefixLength));

        var payload = new byte[bytes.Length - prefixLength];
        Array.Copy(bytes, prefixLength, payload, 0, payload.Length);
        return payload;
    }

    private sealed class WriteOnlyStream : MemoryStream {
        public override bool CanRead => false;
    }

    private sealed class ReadOnlyStream : MemoryStream {
        public override bool CanWrite => false;
    }
}

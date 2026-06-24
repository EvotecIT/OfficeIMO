using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageExtractorTests {
    [Fact]
    public void ExtractPages_RejectsInvalidSelections() {
        byte[] source = BuildThreePagePdf();

        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPages(source, Array.Empty<int>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageExtractor.ExtractPages(source, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageExtractor.ExtractPages(source, 4));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageExtractor.ExtractPageRange(source, 3, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPages((Stream)null!, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPages((Stream)null!, new[] { 1 }));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPages(new MemoryStream(source), (IEnumerable<int>)null!));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPages(new WriteOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPages(source, (Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPages(source, new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPages(new MemoryStream(source), null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPages(new MemoryStream(source), new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRange((Stream)null!, 1, 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRange(new WriteOnlyStream(), 1, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRange(source, null!, 1, 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRange(source, new ReadOnlyStream(), 1, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRange(new MemoryStream(source), null!, 1, 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRange(new MemoryStream(source), new ReadOnlyStream(), 1, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRange(null!, "out.pdf", 1, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRange("input.pdf", (string)null!, 1, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRange("input.pdf", (Stream)null!, 1, 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRange("missing.pdf", new ReadOnlyStream(), 1, 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRange(" ", new MemoryStream(), 1, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRange("input.pdf", (Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRange("missing.pdf", new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRange(" ", "out.pdf", 1, 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRange("input.pdf", " ", 1, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRanges(source, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRanges(source, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageExtractor.ExtractPageRanges(source, PdfPageRange.From(3, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRanges((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRanges(new MemoryStream(source), (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRanges(new WriteOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRanges(source, (Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRanges(source, new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRanges(new MemoryStream(source), null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRanges(new MemoryStream(source), new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRanges("input.pdf", (Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRanges("missing.pdf", new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRanges(" ", new MemoryStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRanges(" ", "out.pdf", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPageRanges("input.pdf", " ", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPages("input.pdf", (Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPages("missing.pdf", new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPages(" ", new MemoryStream(), 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPages(" ", "out.pdf", 1));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.ExtractPages("input.pdf", " ", 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPages((Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPages(new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPages(new MemoryStream(source), null!));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPages(new MemoryStream(source), " "));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPages(null!, "out"));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPages("input.pdf", null!));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPages(" ", "out"));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPages("input.pdf", " "));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPageRanges(source, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPageRanges(source, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageExtractor.SplitPageRanges(source, default(PdfPageRange)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageExtractor.SplitPageRanges(source, PdfPageRange.From(3, 4)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageRange.From(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageRange.From(2, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageRange.Parse(null!));
        Assert.Throws<ArgumentException>(() => PdfPageRange.Parse(" "));
        Assert.Throws<ArgumentException>(() => PdfPageRange.Parse("1-2-3"));
        Assert.Throws<FormatException>(() => PdfPageRange.Parse("two"));
        Assert.Throws<ArgumentException>(() => PdfPageRange.ParseMany("1,,2"));
        Assert.Throws<ArgumentException>(() => PdfPageRange.ParseMany("1; "));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPageRanges((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPageRanges(new MemoryStream(source), (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPageRanges(new WriteOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPageRanges((string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPageRanges(" ", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPageRanges("missing.pdf", Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPageRanges(null!, "out", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPageRanges("input.pdf", null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPageRanges(" ", "out", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPageRanges("input.pdf", " ", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPageRanges("missing.pdf", "out", Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.SplitPageRanges(new MemoryStream(source), null!, "stream", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPageRanges(new MemoryStream(source), " ", "stream", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageExtractor.SplitPageRanges(new WriteOnlyStream(), "out", "stream", Array.Empty<PdfPageRange>()));
    }

    [Fact]
    public void SplitPages_RejectsFileOutputDirectoryBeforeReadingInput() {
        byte[] source = BuildThreePagePdf();
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-split-output-directory-" + Guid.NewGuid().ToString("N"));
        string outputFile = Path.Combine(directory, "not-a-directory");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllText(outputFile, "existing file");

            var pathException = Assert.Throws<ArgumentException>(() =>
                PdfPageExtractor.SplitPages("missing.pdf", outputFile));
            Assert.Equal("outputDirectory", pathException.ParamName);
            Assert.Contains("Output directory refers to a file; a directory path is required.", pathException.Message, StringComparison.Ordinal);

            var streamException = Assert.Throws<ArgumentException>(() =>
                PdfPageExtractor.SplitPages(new MemoryStream(source), outputFile));
            Assert.Equal("outputDirectory", streamException.ParamName);
            Assert.Contains("Output directory refers to a file; a directory path is required.", streamException.Message, StringComparison.Ordinal);

            var rangePathException = Assert.Throws<ArgumentException>(() =>
                PdfPageExtractor.SplitPageRanges("missing.pdf", outputFile, PdfPageRange.From(1, 1)));
            Assert.Equal("outputDirectory", rangePathException.ParamName);
            Assert.Contains("Output directory refers to a file; a directory path is required.", rangePathException.Message, StringComparison.Ordinal);

            var rangeStreamException = Assert.Throws<ArgumentException>(() =>
                PdfPageExtractor.SplitPageRanges(new MemoryStream(source), outputFile, "source", PdfPageRange.From(1, 1)));
            Assert.Equal("outputDirectory", rangeStreamException.ParamName);
            Assert.Contains("Output directory refers to a file; a directory path is required.", rangeStreamException.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractPathOutputs_RejectDirectoryTargets() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-extract-output-path-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputDirectory = Path.Combine(directory, "existing-output");

        try {
            Directory.CreateDirectory(outputDirectory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            var exception = Assert.Throws<ArgumentException>(() =>
                PdfPageExtractor.ExtractPages(inputPath, outputDirectory, 1));

            Assert.Equal("outputPath", exception.ParamName);
            Assert.Contains("Output path refers to a directory; a file path is required.", exception.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}

using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfTextExtractorPageTests {
    [Fact]
    public void ExtractTextByPage_RejectsInvalidInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPage((byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPage((string)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPage((Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText((string)null!, new MemoryStream()));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText("input.pdf", (Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText("input.pdf", (string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllText("input.pdf", " "));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText((Stream)null!, new MemoryStream()));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText(new MemoryStream(), (Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText((byte[])null!, new MemoryStream()));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText(BuildThreePagePdf(), (Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText((Stream)null!, "out.txt"));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText(new MemoryStream(), (string)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText((byte[])null!, "out.txt"));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText(BuildThreePagePdf(), (string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllText(new MemoryStream(), " "));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllText(BuildThreePagePdf(), " "));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown((byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown((string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdown(" "));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown((Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown((string)null!, new MemoryStream()));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown("input.pdf", (Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown("input.pdf", (string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdown("input.pdf", " "));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown((Stream)null!, new MemoryStream()));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown(new MemoryStream(), (Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown((byte[])null!, new MemoryStream()));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown(BuildThreePagePdf(), (Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown((Stream)null!, "out.md"));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown(new MemoryStream(), (string)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown((byte[])null!, "out.md"));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdown(BuildThreePagePdf(), (string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdown(BuildThreePagePdf(), " "));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPage((string)null!, "out"));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPage("input.pdf", (string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractTextByPage("input.pdf", " "));

        using var unreadable = new WriteOnlyStream();
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractTextByPage(unreadable));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdown(unreadable));

        using var readOnlyOutput = new ReadOnlyStream();
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllText("input.pdf", readOnlyOutput));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllText(new MemoryStream(BuildThreePagePdf()), new ReadOnlyStream()));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllText(BuildThreePagePdf(), new ReadOnlyStream()));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdown("input.pdf", new ReadOnlyStream()));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdown(new MemoryStream(BuildThreePagePdf()), new ReadOnlyStream()));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdown(BuildThreePagePdf(), new ReadOnlyStream()));
    }

    [Fact]
    public void ExtractTextByPageRanges_RejectsInvalidInputs() {
        byte[] pdf = BuildThreePagePdf();

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPageRanges(pdf, null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractTextByPageRanges(pdf, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractTextByPageRanges(pdf, default(PdfPageRange)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractTextByPageRanges(pdf, PdfPageRange.From(1, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPageRanges((string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPageRanges((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPageRanges((string)null!, "out", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPageRanges("input.pdf", null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractTextByPageRanges("input.pdf", " ", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges(pdf, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllTextByPageRanges(pdf, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractAllTextByPageRanges(pdf, PdfPageRange.From(1, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges((string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges((string)null!, new MemoryStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges("input.pdf", (Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges((byte[])null!, new MemoryStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges(pdf, (Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges((string)null!, "out.txt", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges("input.pdf", (string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllTextByPageRanges("input.pdf", " ", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges((Stream)null!, "out.txt", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges(new MemoryStream(pdf), (string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges((byte[])null!, "out.txt", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllTextByPageRanges(pdf, (string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdownByPageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdownByPageRanges(pdf, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdownByPageRanges(pdf, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractMarkdownByPageRanges(pdf, default(PdfPageRange)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractMarkdownByPageRanges(pdf, PdfPageRange.From(1, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdownByPageRanges((string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdownByPageRanges((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdownByPageRangesAsDocument((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdownByPageRangesAsDocument(pdf, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdownByPageRangesAsDocument(pdf, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTextExtractor.ExtractMarkdownByPageRangesAsDocument(pdf, PdfPageRange.From(1, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdownByPageRangesAsDocument((string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdownByPageRangesAsDocument((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdownByPageRanges((string)null!, "out", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractMarkdownByPageRanges("input.pdf", (string)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdownByPageRanges("input.pdf", " ", PdfPageRange.From(1, 1)));

        using var unreadable = new WriteOnlyStream();
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractTextByPageRanges(unreadable, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllTextByPageRanges(unreadable, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllTextByPageRanges("input.pdf", new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllTextByPageRanges(new MemoryStream(pdf), new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllTextByPageRanges(pdf, new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdownByPageRanges(unreadable, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractMarkdownByPageRangesAsDocument(unreadable, PdfPageRange.From(1, 1)));
    }

    [Fact]
    public void ExtractTextByPage_RejectsFileOutputDirectoryBeforeReadingInput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-text-output-directory-" + Guid.NewGuid().ToString("N"));
        string outputFile = Path.Combine(directory, "not-a-directory");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllText(outputFile, "existing file");

            var exception = Assert.Throws<ArgumentException>(() =>
                PdfTextExtractor.ExtractTextByPage("missing.pdf", outputFile));

            Assert.Equal("outputDirectory", exception.ParamName);
            Assert.Contains("Output directory refers to a file; a directory path is required.", exception.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractAllText_RejectsInvalidOutputTargetsBeforeReadingInput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-all-text-output-" + Guid.NewGuid().ToString("N"));
        string outputDirectory = Path.Combine(directory, "existing-output");

        try {
            Directory.CreateDirectory(outputDirectory);

            var pathException = Assert.Throws<ArgumentException>(() =>
                PdfTextExtractor.ExtractAllText("missing.pdf", outputDirectory));

            Assert.Equal("outputPath", pathException.ParamName);
            Assert.Contains("Output path refers to a directory; a file path is required.", pathException.Message, StringComparison.Ordinal);

            var streamException = Assert.Throws<ArgumentException>(() =>
                PdfTextExtractor.ExtractAllText("missing.pdf", new ReadOnlyStream()));

            Assert.Equal("outputStream", streamException.ParamName);
            Assert.Contains("Stream must be writable.", streamException.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}

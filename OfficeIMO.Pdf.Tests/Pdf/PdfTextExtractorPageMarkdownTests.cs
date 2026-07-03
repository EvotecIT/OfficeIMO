using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfTextExtractorPageTests {
    [Fact]
    public void ExtractMarkdown_WritesLogicalMarkdownToPathAndStreamsForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-markdown-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputPath = Path.Combine(directory, "markdown", "all.md");
        var options = new PdfTextLayoutOptions {
            ForceSingleColumn = true
        };

        try {
            Directory.CreateDirectory(directory);
            byte[] pdf = BuildMarkdownPdf();
            File.WriteAllBytes(inputPath, pdf);

            string markdown = PdfTextExtractor.ExtractMarkdown(pdf, options);

            Assert.Contains("# Markdown Heading", markdown, StringComparison.Ordinal);
            Assert.Contains("Markdownreadbackmarker.", Normalize(markdown), StringComparison.Ordinal);
            Assert.Contains("| Code | Name | Qty |", markdown, StringComparison.Ordinal);
            Assert.Contains("| A-100 | Alpha | 2 |", markdown, StringComparison.Ordinal);

            PdfTextExtractor.ExtractMarkdown(inputPath, outputPath, options);
            Assert.True(File.Exists(outputPath));
            Assert.Contains("# Markdown Heading", File.ReadAllText(outputPath, Encoding.UTF8), StringComparison.Ordinal);

            using var pathOutput = CreateOutputStream(out int pathPrefixLength);
            PdfTextExtractor.ExtractMarkdown(inputPath, pathOutput, options);
            Assert.Contains("| A-100 | Alpha | 2 |", GetOutputText(pathOutput, pathPrefixLength), StringComparison.Ordinal);

            using var streamInput = BuildPrefixedStream(pdf);
            streamInput.Position = 5;
            using var streamOutput = CreateOutputStream(out int streamPrefixLength);
            PdfTextExtractor.ExtractMarkdown(streamInput, streamOutput, options);
            Assert.Contains("Markdownreadbackmarker.", Normalize(GetOutputText(streamOutput, streamPrefixLength)), StringComparison.Ordinal);

            using var byteOutput = CreateOutputStream(out int bytePrefixLength);
            PdfTextExtractor.ExtractMarkdown(pdf, byteOutput, options);
            Assert.Contains("# Markdown Heading", GetOutputText(byteOutput, bytePrefixLength), StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractMarkdownByPageRanges_ReturnsCallerOrderAndWritesMarkdownFiles() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-markdown-ranges-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        var options = new PdfTextLayoutOptions {
            ForceSingleColumn = true
        };

        try {
            Directory.CreateDirectory(directory);
            byte[] pdf = BuildThreePageMarkdownPdf();
            File.WriteAllBytes(inputPath, pdf);

            IReadOnlyList<string> pages = PdfTextExtractor.ExtractMarkdownByPageRanges(
                pdf,
                options,
                null,
                PdfPageRange.ParseMany("3,1-2,2"));

            Assert.Equal(4, pages.Count);
            Assert.Contains("# Third Page", pages[0], StringComparison.Ordinal);
            Assert.Contains("# First Page", pages[1], StringComparison.Ordinal);
            Assert.Contains("# Second Page", pages[2], StringComparison.Ordinal);
            Assert.Contains("# Second Page", pages[3], StringComparison.Ordinal);

            string selectedDocument = PdfTextExtractor.ExtractMarkdownByPageRangesAsDocument(
                pdf,
                options,
                new PdfLogicalMarkdownOptions {
                    PageSeparator = "***"
                },
                PdfPageRange.ParseMany("2,1"));

            AssertContainsInOrder(
                Normalize(selectedDocument),
                "#SecondPage",
                "***",
                "#FirstPage");

            string outputDirectory = Path.Combine(directory, "path-markdown");
            IReadOnlyList<string> paths = PdfTextExtractor.ExtractMarkdownByPageRanges(inputPath, outputDirectory, options, null, PdfPageRange.ParseMany("3,1-2,2"));

            Assert.Equal(4, paths.Count);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0003.md"), paths[0]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0001.md"), paths[1]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0002.md"), paths[2]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0002-occurrence-0002.md"), paths[3]);
            Assert.Contains("# Third Page", File.ReadAllText(paths[0], Encoding.UTF8), StringComparison.Ordinal);

            using var stream = BuildPrefixedStream(pdf);
            stream.Position = 5;
            string streamOutputDirectory = Path.Combine(directory, "stream-markdown");
            IReadOnlyList<string> streamPaths = PdfTextExtractor.ExtractMarkdownByPageRanges(
                stream,
                streamOutputDirectory,
                "stream-source.pdf",
                options,
                null,
                PdfPageRange.ParseMany("2-3"));

            Assert.Equal(2, streamPaths.Count);
            Assert.Equal(Path.Combine(streamOutputDirectory, "stream-source-page-0002.md"), streamPaths[0]);
            Assert.Equal(Path.Combine(streamOutputDirectory, "stream-source-page-0003.md"), streamPaths[1]);
            Assert.Contains("# Second Page", File.ReadAllText(streamPaths[0], Encoding.UTF8), StringComparison.Ordinal);

            string byteOutputDirectory = Path.Combine(directory, "byte-markdown");
            IReadOnlyList<string> bytePaths = PdfTextExtractor.ExtractMarkdownByPage(pdf, byteOutputDirectory, "byte-source.pdf", options);

            Assert.Equal(3, bytePaths.Count);
            Assert.Equal(Path.Combine(byteOutputDirectory, "byte-source-page-0001.md"), bytePaths[0]);
            Assert.Equal(Path.Combine(byteOutputDirectory, "byte-source-page-0002.md"), bytePaths[1]);
            Assert.Equal(Path.Combine(byteOutputDirectory, "byte-source-page-0003.md"), bytePaths[2]);
            Assert.Contains("# First Page", File.ReadAllText(bytePaths[0], Encoding.UTF8), StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}

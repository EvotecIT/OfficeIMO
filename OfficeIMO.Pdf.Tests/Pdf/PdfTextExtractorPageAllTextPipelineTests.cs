using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfTextExtractorPageTests {
    [Fact]
    public void ExtractAllText_WritesTextToPathAndOutputStreamsForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-all-text-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputPath = Path.Combine(directory, "text", "all.txt");

        try {
            Directory.CreateDirectory(directory);
            byte[] pdf = BuildThreePagePdf();
            File.WriteAllBytes(inputPath, pdf);

            PdfTextExtractor.ExtractAllText(inputPath, outputPath);

            Assert.True(File.Exists(outputPath));
            string fileText = Normalize(File.ReadAllText(outputPath, Encoding.UTF8));
            Assert.Contains("Firstpagemarker", fileText, StringComparison.Ordinal);
            Assert.Contains("Secondpagemarker", fileText, StringComparison.Ordinal);
            Assert.Contains("Thirdpagemarker", fileText, StringComparison.Ordinal);

            using var pathOutput = CreateOutputStream(out int pathPrefixLength);
            PdfTextExtractor.ExtractAllText(inputPath, pathOutput);
            string pathOutputText = Normalize(GetOutputText(pathOutput, pathPrefixLength));
            Assert.Contains("Secondpagemarker", pathOutputText, StringComparison.Ordinal);

            using var streamInput = BuildPrefixedStream(pdf);
            streamInput.Position = 5;
            using var streamOutput = CreateOutputStream(out int streamPrefixLength);
            PdfTextExtractor.ExtractAllText(streamInput, streamOutput);
            string streamOutputText = Normalize(GetOutputText(streamOutput, streamPrefixLength));
            Assert.Contains("Thirdpagemarker", streamOutputText, StringComparison.Ordinal);

            using var byteOutput = CreateOutputStream(out int bytePrefixLength);
            PdfTextExtractor.ExtractAllText(pdf, byteOutput);
            string byteOutputText = Normalize(GetOutputText(byteOutput, bytePrefixLength));
            Assert.Contains("Firstpagemarker", byteOutputText, StringComparison.Ordinal);

            string columnOutputPath = Path.Combine(directory, "text", "columns.txt");
            PdfTextExtractor.ExtractAllText(inputPath, columnOutputPath, new PdfTextLayoutOptions {
                MarginLeft = 36,
                MarginRight = 36,
                MinGutterWidth = 24
            });

            string columnText = Normalize(File.ReadAllText(columnOutputPath, Encoding.UTF8));
            Assert.Contains("Firstpagemarker", columnText, StringComparison.Ordinal);

            using var streamInputForPath = BuildPrefixedStream(pdf);
            streamInputForPath.Position = 5;
            string streamPathOutput = Path.Combine(directory, "text", "stream-all.txt");
            PdfTextExtractor.ExtractAllText(streamInputForPath, streamPathOutput);
            string streamPathText = Normalize(File.ReadAllText(streamPathOutput, Encoding.UTF8));
            Assert.Contains("Secondpagemarker", streamPathText, StringComparison.Ordinal);

            string bytePathOutput = Path.Combine(directory, "text", "byte-all.txt");
            PdfTextExtractor.ExtractAllText(pdf, bytePathOutput);
            string bytePathText = Normalize(File.ReadAllText(bytePathOutput, Encoding.UTF8));
            Assert.Contains("Thirdpagemarker", bytePathText, StringComparison.Ordinal);

            string byteColumnOutput = Path.Combine(directory, "text", "byte-columns.txt");
            PdfTextExtractor.ExtractAllText(BuildTwoColumnPdf(), byteColumnOutput, new PdfTextLayoutOptions {
                MarginLeft = 36,
                MarginRight = 36,
                MinGutterWidth = 24
            });
            AssertColumnAwareTextOrder(File.ReadAllText(byteColumnOutput, Encoding.UTF8));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractAllTextByPageRanges_ConcatenatesSelectedPagesForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-all-text-page-ranges-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputPath = Path.Combine(directory, "text", "selected.txt");

        try {
            Directory.CreateDirectory(directory);
            byte[] pdf = BuildThreePagePdf();
            File.WriteAllBytes(inputPath, pdf);

            string text = PdfTextExtractor.ExtractAllTextByPageRanges(pdf, PdfPageRange.ParseMany("3,1-2,2"));
            string normalized = Normalize(text);

            AssertContainsInOrder(normalized, "Thirdpagemarker", "Firstpagemarker", "Secondpagemarker", "Secondpagemarker");
            Assert.Equal(2, CountOccurrences(normalized, "Secondpagemarker"));

            PdfTextExtractor.ExtractAllTextByPageRanges(inputPath, outputPath, PdfPageRange.ParseMany("2-3"));
            string fileText = Normalize(File.ReadAllText(outputPath, Encoding.UTF8));
            Assert.DoesNotContain("Firstpagemarker", fileText, StringComparison.Ordinal);
            AssertContainsInOrder(fileText, "Secondpagemarker", "Thirdpagemarker");

            using var streamInput = BuildPrefixedStream(pdf);
            streamInput.Position = 5;
            using var streamOutput = CreateOutputStream(out int prefixLength);
            PdfTextExtractor.ExtractAllTextByPageRanges(streamInput, streamOutput, PdfPageRange.ParseMany("1,3"));
            string streamText = Normalize(GetOutputText(streamOutput, prefixLength));
            AssertContainsInOrder(streamText, "Firstpagemarker", "Thirdpagemarker");
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractAllTextByPageRanges_WithLayoutOptionsUsesColumnAwareReadingOrder() {
        var options = new PdfTextLayoutOptions {
            MarginLeft = 36,
            MarginRight = 36,
            MinGutterWidth = 24
        };

        string text = PdfTextExtractor.ExtractAllTextByPageRanges(
            BuildTwoColumnPdf(),
            options,
            PdfPageRange.From(1, 1));

        AssertColumnAwareTextOrder(text);
    }
}

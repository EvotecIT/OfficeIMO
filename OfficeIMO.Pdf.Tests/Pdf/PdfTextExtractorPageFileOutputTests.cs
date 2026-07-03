using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfTextExtractorPageTests {
    [Fact]
    public void ExtractTextByPage_WritesTextFilesToDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-text-pages-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputDirectory = Path.Combine(directory, "text");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            IReadOnlyList<string> paths = PdfTextExtractor.ExtractTextByPage(inputPath, outputDirectory);

            Assert.Equal(3, paths.Count);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0001.txt"), paths[0]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0002.txt"), paths[1]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0003.txt"), paths[2]);

            Assert.Contains("Firstpagemarker", Normalize(File.ReadAllText(paths[0], Encoding.UTF8)), StringComparison.Ordinal);
            Assert.Contains("Secondpagemarker", Normalize(File.ReadAllText(paths[1], Encoding.UTF8)), StringComparison.Ordinal);
            Assert.Contains("Thirdpagemarker", Normalize(File.ReadAllText(paths[2], Encoding.UTF8)), StringComparison.Ordinal);

            using var stream = new MemoryStream(BuildThreePagePdf());
            string streamOutputDirectory = Path.Combine(directory, "stream-text");
            IReadOnlyList<string> streamPaths = PdfTextExtractor.ExtractTextByPage(stream, streamOutputDirectory, "stream-source.pdf");

            Assert.Equal(3, streamPaths.Count);
            Assert.Equal(Path.Combine(streamOutputDirectory, "stream-source-page-0001.txt"), streamPaths[0]);
            Assert.Equal(Path.Combine(streamOutputDirectory, "stream-source-page-0002.txt"), streamPaths[1]);
            Assert.Equal(Path.Combine(streamOutputDirectory, "stream-source-page-0003.txt"), streamPaths[2]);
            Assert.Contains("Firstpagemarker", Normalize(File.ReadAllText(streamPaths[0], Encoding.UTF8)), StringComparison.Ordinal);

            string byteOutputDirectory = Path.Combine(directory, "byte-text");
            IReadOnlyList<string> bytePaths = PdfTextExtractor.ExtractTextByPage(BuildThreePagePdf(), byteOutputDirectory, "byte-source.pdf");

            Assert.Equal(3, bytePaths.Count);
            Assert.Equal(Path.Combine(byteOutputDirectory, "byte-source-page-0001.txt"), bytePaths[0]);
            Assert.Equal(Path.Combine(byteOutputDirectory, "byte-source-page-0002.txt"), bytePaths[1]);
            Assert.Equal(Path.Combine(byteOutputDirectory, "byte-source-page-0003.txt"), bytePaths[2]);
            Assert.Contains("Secondpagemarker", Normalize(File.ReadAllText(bytePaths[1], Encoding.UTF8)), StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractTextByPageRanges_WritesSelectedSourcePageFilesToDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-text-page-ranges-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputDirectory = Path.Combine(directory, "text");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            IReadOnlyList<string> paths = PdfTextExtractor.ExtractTextByPageRanges(inputPath, outputDirectory, PdfPageRange.ParseMany("3,1-2,2"));

            Assert.Equal(4, paths.Count);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0003.txt"), paths[0]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0001.txt"), paths[1]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0002.txt"), paths[2]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0002-occurrence-0002.txt"), paths[3]);
            Assert.NotEqual(paths[2], paths[3]);

            Assert.Contains("Thirdpagemarker", Normalize(File.ReadAllText(paths[0], Encoding.UTF8)), StringComparison.Ordinal);
            Assert.Contains("Firstpagemarker", Normalize(File.ReadAllText(paths[1], Encoding.UTF8)), StringComparison.Ordinal);
            Assert.Contains("Secondpagemarker", Normalize(File.ReadAllText(paths[2], Encoding.UTF8)), StringComparison.Ordinal);
            Assert.Contains("Secondpagemarker", Normalize(File.ReadAllText(paths[3], Encoding.UTF8)), StringComparison.Ordinal);

            using var stream = new MemoryStream(BuildThreePagePdf());
            string streamOutputDirectory = Path.Combine(directory, "stream-text");
            IReadOnlyList<string> streamPaths = PdfTextExtractor.ExtractTextByPageRanges(stream, streamOutputDirectory, "stream-source.pdf", PdfPageRange.ParseMany("2-3"));

            Assert.Equal(2, streamPaths.Count);
            Assert.Equal(Path.Combine(streamOutputDirectory, "stream-source-page-0002.txt"), streamPaths[0]);
            Assert.Equal(Path.Combine(streamOutputDirectory, "stream-source-page-0003.txt"), streamPaths[1]);
            Assert.Contains("Thirdpagemarker", Normalize(File.ReadAllText(streamPaths[1], Encoding.UTF8)), StringComparison.Ordinal);

            string byteOutputDirectory = Path.Combine(directory, "byte-text");
            IReadOnlyList<string> bytePaths = PdfTextExtractor.ExtractTextByPageRanges(BuildThreePagePdf(), byteOutputDirectory, "byte-source.pdf", PdfPageRange.ParseMany("1,3"));

            Assert.Equal(2, bytePaths.Count);
            Assert.Equal(Path.Combine(byteOutputDirectory, "byte-source-page-0001.txt"), bytePaths[0]);
            Assert.Equal(Path.Combine(byteOutputDirectory, "byte-source-page-0003.txt"), bytePaths[1]);
            Assert.Contains("Firstpagemarker", Normalize(File.ReadAllText(bytePaths[0], Encoding.UTF8)), StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractTextByPageRanges_WithLayoutOptionsWritesSelectedSourcePageFiles() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-text-page-range-options-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "columns.pdf");
        var options = new PdfTextLayoutOptions {
            MarginLeft = 36,
            MarginRight = 36,
            MinGutterWidth = 24
        };

        try {
            Directory.CreateDirectory(directory);
            byte[] pdf = BuildTwoColumnPdf();
            File.WriteAllBytes(inputPath, pdf);

            string pathOutputDirectory = Path.Combine(directory, "path-text");
            IReadOnlyList<string> pathFiles = PdfTextExtractor.ExtractTextByPageRanges(
                inputPath,
                pathOutputDirectory,
                options,
                PdfPageRange.From(1, 1));

            string pathFile = Assert.Single(pathFiles);
            Assert.Equal(Path.Combine(pathOutputDirectory, "columns-page-0001.txt"), pathFile);
            AssertColumnAwareTextOrder(File.ReadAllText(pathFile, Encoding.UTF8));

            using var stream = BuildPrefixedStream(pdf);
            stream.Position = 5;
            string streamOutputDirectory = Path.Combine(directory, "stream-text");
            IReadOnlyList<string> streamFiles = PdfTextExtractor.ExtractTextByPageRanges(
                stream,
                streamOutputDirectory,
                "stream-columns.pdf",
                options,
                PdfPageRange.From(1, 1));

            string streamFile = Assert.Single(streamFiles);
            Assert.Equal(Path.Combine(streamOutputDirectory, "stream-columns-page-0001.txt"), streamFile);
            AssertColumnAwareTextOrder(File.ReadAllText(streamFile, Encoding.UTF8));

            string byteOutputDirectory = Path.Combine(directory, "byte-text");
            IReadOnlyList<string> byteFiles = PdfTextExtractor.ExtractTextByPageRanges(
                pdf,
                byteOutputDirectory,
                "byte-columns.pdf",
                options,
                PdfPageRange.From(1, 1));

            string byteFile = Assert.Single(byteFiles);
            Assert.Equal(Path.Combine(byteOutputDirectory, "byte-columns-page-0001.txt"), byteFile);
            AssertColumnAwareTextOrder(File.ReadAllText(byteFile, Encoding.UTF8));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}

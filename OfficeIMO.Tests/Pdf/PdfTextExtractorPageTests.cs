using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTextExtractorPageTests {
    [Fact]
    public void ExtractTextByPage_ReturnsOneTextEntryPerPage() {
        byte[] pdf = BuildThreePagePdf();

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPage(pdf);

        Assert.Equal(3, pages.Count);
        Assert.Contains("Firstpagemarker", Normalize(pages[0]), StringComparison.Ordinal);
        Assert.Contains("Secondpagemarker", Normalize(pages[1]), StringComparison.Ordinal);
        Assert.Contains("Thirdpagemarker", Normalize(pages[2]), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPageRanges_ReturnsParsedRangesInCallerOrderWithRepeatedPages() {
        byte[] pdf = BuildThreePagePdf();

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPageRanges(pdf, PdfPageRange.ParseMany("3,1-2,2"));

        Assert.Equal(4, pages.Count);
        Assert.Contains("Thirdpagemarker", Normalize(pages[0]), StringComparison.Ordinal);
        Assert.Contains("Firstpagemarker", Normalize(pages[1]), StringComparison.Ordinal);
        Assert.Contains("Secondpagemarker", Normalize(pages[2]), StringComparison.Ordinal);
        Assert.Contains("Secondpagemarker", Normalize(pages[3]), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_ReadsFromCurrentStreamPosition() {
        byte[] pdf = BuildThreePagePdf();
        using var stream = BuildPrefixedStream(pdf);
        stream.Position = 5;

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPage(stream);

        Assert.Equal(3, pages.Count);
        Assert.Contains("Secondpagemarker", Normalize(pages[1]), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPageRanges_ReadsFromCurrentStreamPosition() {
        byte[] pdf = BuildThreePagePdf();
        using var stream = BuildPrefixedStream(pdf);
        stream.Position = 5;

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPageRanges(stream, PdfPageRange.ParseMany("2-3"));

        Assert.Equal(2, pages.Count);
        Assert.Contains("Secondpagemarker", Normalize(pages[0]), StringComparison.Ordinal);
        Assert.Contains("Thirdpagemarker", Normalize(pages[1]), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_WithLayoutOptionsUsesColumnAwareReadingOrder() {
        byte[] pdf = BuildTwoColumnPdf();

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPage(pdf, new PdfTextLayoutOptions {
            MarginLeft = 36,
            MarginRight = 36,
            MinGutterWidth = 24
        });

        string text = Normalize(pages[0]);
        int leftStart = text.IndexOf("LeftStart", StringComparison.Ordinal);
        int leftFinish = text.IndexOf("LeftFinish", StringComparison.Ordinal);
        int rightStart = text.IndexOf("RightStart", StringComparison.Ordinal);
        int rightFinish = text.IndexOf("RightFinish", StringComparison.Ordinal);

        Assert.Single(pages);
        Assert.True(leftStart >= 0, "Expected left column start marker to be extracted.");
        Assert.True(leftFinish > leftStart, "Expected left column markers to preserve top-to-bottom order.");
        Assert.True(rightStart >= 0, "Expected right column start marker to be extracted.");
        Assert.True(rightFinish > rightStart, "Expected right column markers to preserve top-to-bottom order.");
        Assert.True(leftFinish < rightStart,
            $"Expected column-aware extraction to finish the left column before reading the right column. Text: {pages[0]}");
    }

    [Fact]
    public void GetTextSpans_UsesStandardFontMetricsWhenWidthsAreOmitted() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.TimesRoman,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("WWWW"))
            .ToBytes();

        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Load(pdf)
                .Pages[0]
                .GetTextSpans(),
            item => item.Text == "WWWW");

        Assert.Equal(37.76, span.Advance, 2);
    }

    [Fact]
    public void GetTextSpans_UsesWinAnsiPunctuationMetricsWhenWidthsAreOmitted() {
        const string text = "\u201CWait\u201D\u2014ok\u2026";
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.TimesRoman,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text(text))
            .ToBytes();

        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Load(pdf)
                .Pages[0]
                .GetTextSpans(),
            item => item.Text == text);

        Assert.Equal(58.32, span.Advance, 2);
    }

    [Fact]
    public void GetTextSpans_UsesWinAnsiAccentedLetterMetricsWhenWidthsAreOmitted() {
        const string text = "r\u00E9sum\u00E9";
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.TimesRoman,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text(text))
            .ToBytes();

        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Load(pdf)
                .Pages[0]
                .GetTextSpans(),
            item => item.Text == text);

        Assert.Equal(28.88, span.Advance, 2);
    }

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

    private static byte[] BuildThreePagePdf() {
        var doc = PdfDoc.Create();
        doc.Compose(compose => {
            compose.Page(page => {
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("First page marker"))));
            });

            compose.Page(page => {
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Second page marker"))));
            });

            compose.Page(page => {
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Third page marker"))));
            });
        });

        return doc.ToBytes();
    }

    private static byte[] BuildTwoColumnPdf() {
        return PdfDoc.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row => row
                            .Gap(36)
                            .Column(50, column => column
                                .Paragraph(p => p.Text("Left Start marker"))
                                .Paragraph(p => p.Text("Left Finish marker")))
                            .Column(50, column => column
                                .Paragraph(p => p.Text("Right Start marker"))
                                .Paragraph(p => p.Text("Right Finish marker")))))))
            .ToBytes();
    }

    private static byte[] BuildMarkdownPdf() {
        return PdfDoc.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .H1("Markdown Heading")
            .Paragraph(p => p.Text("Markdown readback marker."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
    }

    private static byte[] BuildThreePageMarkdownPdf() {
        return PdfDoc.Create(new PdfOptions {
                PageWidth = 300,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 10
            })
            .H1("First Page")
            .Paragraph(p => p.Text("First markdown marker."))
            .PageBreak()
            .H1("Second Page")
            .Paragraph(p => p.Text("Second markdown marker."))
            .PageBreak()
            .H1("Third Page")
            .Paragraph(p => p.Text("Third markdown marker."))
            .ToBytes();
    }

    private static MemoryStream BuildPrefixedStream(byte[] pdf) {
        var data = new byte[pdf.Length + 5];
        data[0] = 1;
        data[1] = 2;
        data[2] = 3;
        data[3] = 4;
        data[4] = 5;
        Array.Copy(pdf, 0, data, 5, pdf.Length);
        return new MemoryStream(data);
    }

    private static MemoryStream CreateOutputStream(out int prefixLength) {
        byte[] prefix = Encoding.ASCII.GetBytes("output-prefix");
        var stream = new MemoryStream();
        stream.Write(prefix, 0, prefix.Length);
        prefixLength = prefix.Length;
        return stream;
    }

    private static string GetOutputText(MemoryStream output, int prefixLength) {
        byte[] bytes = output.ToArray();
        Assert.True(bytes.Length > prefixLength);
        Assert.Equal("output-prefix", Encoding.ASCII.GetString(bytes, 0, prefixLength));
        return Encoding.UTF8.GetString(bytes, prefixLength, bytes.Length - prefixLength);
    }

    private static string Normalize(string text) {
        return text.Replace(" ", string.Empty);
    }

    private static void AssertColumnAwareTextOrder(string text) {
        string normalized = Normalize(text);
        int leftStart = normalized.IndexOf("LeftStart", StringComparison.Ordinal);
        int leftFinish = normalized.IndexOf("LeftFinish", StringComparison.Ordinal);
        int rightStart = normalized.IndexOf("RightStart", StringComparison.Ordinal);
        int rightFinish = normalized.IndexOf("RightFinish", StringComparison.Ordinal);

        Assert.True(leftStart >= 0, "Expected left column start marker to be extracted.");
        Assert.True(leftFinish > leftStart, "Expected left column markers to preserve top-to-bottom order.");
        Assert.True(rightStart >= 0, "Expected right column start marker to be extracted.");
        Assert.True(rightFinish > rightStart, "Expected right column markers to preserve top-to-bottom order.");
        Assert.True(leftFinish < rightStart,
            $"Expected column-aware extraction to finish the left column before reading the right column. Text: {text}");
    }

    private static void AssertContainsInOrder(string text, params string[] markers) {
        int lastIndex = -1;
        for (int i = 0; i < markers.Length; i++) {
            int index = text.IndexOf(markers[i], lastIndex + 1, StringComparison.Ordinal);
            Assert.True(index > lastIndex, $"Expected marker '{markers[i]}' after index {lastIndex}. Text: {text}");
            lastIndex = index;
        }
    }

    private static int CountOccurrences(string text, string marker) {
        int count = 0;
        int index = 0;
        while (true) {
            index = text.IndexOf(marker, index, StringComparison.Ordinal);
            if (index < 0) {
                return count;
            }

            count++;
            index += marker.Length;
        }
    }

    private sealed class WriteOnlyStream : Stream {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => 0;

        public override long Position {
            get => 0;
            set => throw new NotSupportedException();
        }

        public override void Flush() {
        }

        public override int Read(byte[] buffer, int offset, int count) {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin) {
            throw new NotSupportedException();
        }

        public override void SetLength(long value) {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count) {
        }
    }

    private sealed class ReadOnlyStream : MemoryStream {
        public override bool CanWrite => false;
    }
}

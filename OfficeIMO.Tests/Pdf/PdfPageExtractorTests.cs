using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageExtractorTests {
    [Fact]
    public void ExtractPages_CopiesSelectedPagesInRequestedOrder() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 3, 1);

        using var pdf = PdfDocument.Open(new MemoryStream(extracted));
        Assert.Equal(2, pdf.NumberOfPages);

        var read = PdfReadDocument.Load(extracted);
        Assert.Equal(2, read.Pages.Count);

        string text = NormalizeExtractedText(read.ExtractText());
        Assert.Contains("Thirdpagemarker", text);
        Assert.Contains("Firstpagemarker", text);
        Assert.DoesNotContain("Secondpagemarker", text);
        Assert.True(text.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < text.IndexOf("Firstpagemarker", StringComparison.Ordinal));

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal("Extraction sample", info.Metadata.Title);
        Assert.Equal("OfficeIMO", info.Metadata.Author);
        Assert.Equal(2, info.PageCount);
        Assert.Equal(300, info.Pages[0].Width);
        Assert.Equal(500, info.Pages[0].Height);
        Assert.Equal(595, info.Pages[1].Width);
        Assert.Equal(842, info.Pages[1].Height);
    }

    [Fact]
    public void ExtractPages_ClonesDuplicateSelectionsInRequestedOrder() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 3, 3, 1);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal(3, info.PageCount);
        Assert.Equal(300, info.Pages[0].Width);
        Assert.Equal(500, info.Pages[0].Height);
        Assert.Equal(300, info.Pages[1].Width);
        Assert.Equal(500, info.Pages[1].Height);
        Assert.Equal(595, info.Pages[2].Width);
        Assert.Equal(842, info.Pages[2].Height);

        string text = NormalizeExtractedText(PdfReadDocument.Load(extracted).ExtractText());
        Assert.Equal(2, CountOccurrences(text, "Thirdpagemarker"));
        Assert.Equal(1, CountOccurrences(text, "Firstpagemarker"));
        Assert.DoesNotContain("Secondpagemarker", text);
        AssertContainsInOrder(text, "Thirdpagemarker", "Thirdpagemarker", "Firstpagemarker");
    }

    [Fact]
    public void ExtractPageRange_CopiesInclusiveRange() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPageRange(source, 2, 3);

        var read = PdfReadDocument.Load(extracted);
        string text = NormalizeExtractedText(read.ExtractText());

        Assert.Equal(2, read.Pages.Count);
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }

    [Fact]
    public void ExtractPageRange_AcceptsPdfPageRange() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPageRange(source, PdfPageRange.From(2, 3));

        var read = PdfReadDocument.Load(extracted);
        string text = NormalizeExtractedText(read.ExtractText());

        Assert.Equal(2, read.Pages.Count);
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }

    [Fact]
    public void ExtractPageRanges_CombinesParsedRangesInCallerOrder() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPageRanges(source, PdfPageRange.ParseMany("3,1-2"));

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal(3, info.PageCount);
        Assert.Equal(300, info.Pages[0].Width);
        Assert.Equal(500, info.Pages[0].Height);
        Assert.Equal(595, info.Pages[1].Width);
        Assert.Equal(842, info.Pages[1].Height);
        Assert.Equal(612, info.Pages[2].Width);
        Assert.Equal(792, info.Pages[2].Height);

        string text = NormalizeExtractedText(PdfReadDocument.Load(extracted).ExtractText());
        AssertContainsInOrder(text, "Thirdpagemarker", "Firstpagemarker", "Secondpagemarker");
    }

    [Fact]
    public void ExtractPageRanges_PreservesDuplicateAndOverlappingRanges() {
        byte[] source = BuildThreePagePdf();

        byte[] extracted = PdfPageExtractor.ExtractPageRanges(source, PdfPageRange.ParseMany("2,2-3"));

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal(3, info.PageCount);

        string text = NormalizeExtractedText(PdfReadDocument.Load(extracted).ExtractText());
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.Equal(2, CountOccurrences(text, "Secondpagemarker"));
        Assert.Equal(1, CountOccurrences(text, "Thirdpagemarker"));
        AssertContainsInOrder(text, "Secondpagemarker", "Secondpagemarker", "Thirdpagemarker");
    }

    [Fact]
    public void PdfPageRange_ParseReadsSingleAndInclusiveRanges() {
        Assert.Equal(PdfPageRange.From(3, 3), PdfPageRange.Parse("3"));
        Assert.Equal(PdfPageRange.From(1, 3), PdfPageRange.Parse(" 1-3 "));
        Assert.Equal(PdfPageRange.From(2, 4), PdfPageRange.Parse("2 .. 4"));
        Assert.Equal("2-4", PdfPageRange.Parse("2..4").ToString());
    }

    [Fact]
    public void PdfPageRange_ParseManyReadsWrapperFriendlyLists() {
        PdfPageRange[] ranges = PdfPageRange.ParseMany("1-2, 3; 2..3");

        Assert.Equal(new[] {
            PdfPageRange.From(1, 2),
            PdfPageRange.From(3, 3),
            PdfPageRange.From(2, 3)
        }, ranges);
    }

    [Fact]
    public void PdfPageRange_TryParseRejectsInvalidText() {
        Assert.True(PdfPageRange.TryParse("2-3", out PdfPageRange singleRange));
        Assert.Equal(PdfPageRange.From(2, 3), singleRange);

        Assert.True(PdfPageRange.TryParseMany("1,3-4", out PdfPageRange[] ranges));
        Assert.Equal(new[] { PdfPageRange.From(1, 1), PdfPageRange.From(3, 4) }, ranges);

        Assert.False(PdfPageRange.TryParse(null, out _));
        Assert.False(PdfPageRange.TryParse(" ", out _));
        Assert.False(PdfPageRange.TryParse("0", out _));
        Assert.False(PdfPageRange.TryParse("4-2", out _));
        Assert.False(PdfPageRange.TryParse("1-2-3", out _));
        Assert.False(PdfPageRange.TryParse("two", out _));
        Assert.False(PdfPageRange.TryParseMany(null, out _));
        Assert.False(PdfPageRange.TryParseMany("1,,2", out _));
    }

    [Fact]
    public void SplitPageRanges_AcceptsParsedRangeText() {
        byte[] source = BuildThreePagePdf();

        IReadOnlyList<byte[]> ranges = PdfPageExtractor.SplitPageRanges(source, PdfPageRange.ParseMany("1-2,3"));

        Assert.Equal(2, ranges.Count);
        string firstText = NormalizeExtractedText(PdfReadDocument.Load(ranges[0]).ExtractText());
        Assert.Contains("Firstpagemarker", firstText);
        Assert.Contains("Secondpagemarker", firstText);
        Assert.DoesNotContain("Thirdpagemarker", firstText);

        string secondText = NormalizeExtractedText(PdfReadDocument.Load(ranges[1]).ExtractText());
        Assert.DoesNotContain("Firstpagemarker", secondText);
        Assert.DoesNotContain("Secondpagemarker", secondText);
        Assert.Contains("Thirdpagemarker", secondText);
    }

    [Fact]
    public void ExtractPages_ReadsFromCurrentStreamPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        byte[] extracted = PdfPageExtractor.ExtractPages(stream, 2);

        var read = PdfReadDocument.Load(extracted);
        string text = NormalizeExtractedText(read.ExtractText());
        Assert.Single(read.Pages);
        Assert.Contains("Secondpagemarker", text);
        Assert.DoesNotContain("Firstpagemarker", text);
    }

    [Fact]
    public void ExtractPageRange_ReadsFromCurrentStreamPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        byte[] extracted = PdfPageExtractor.ExtractPageRange(stream, 1, 2);

        string text = NormalizeExtractedText(PdfReadDocument.Load(extracted).ExtractText());
        Assert.Contains("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.DoesNotContain("Thirdpagemarker", text);
    }

    [Fact]
    public void ExtractPageRanges_ReadsFromCurrentStreamPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        byte[] extracted = PdfPageExtractor.ExtractPageRanges(stream, PdfPageRange.From(3, 3), PdfPageRange.From(1, 1));

        string text = NormalizeExtractedText(PdfReadDocument.Load(extracted).ExtractText());
        AssertContainsInOrder(text, "Thirdpagemarker", "Firstpagemarker");
        Assert.DoesNotContain("Secondpagemarker", text);
    }

    [Fact]
    public void ExtractPages_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageExtractor.ExtractPages(source, output, 3);

        byte[] extracted = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        string text = NormalizeExtractedText(PdfReadDocument.Load(extracted).ExtractText());
        Assert.Contains("Thirdpagemarker", text);
        Assert.DoesNotContain("Firstpagemarker", text);
    }

    [Fact]
    public void ExtractPageRanges_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageExtractor.ExtractPageRanges(source, output, PdfPageRange.From(2, 2), PdfPageRange.From(1, 1));

        byte[] extracted = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        string text = NormalizeExtractedText(PdfReadDocument.Load(extracted).ExtractText());
        AssertContainsInOrder(text, "Secondpagemarker", "Firstpagemarker");
        Assert.DoesNotContain("Thirdpagemarker", text);
    }

    [Fact]
    public void ExtractPages_WritesStreamInputToOutputStream() {
        byte[] source = BuildThreePagePdf();
        byte[] inputPrefix = Encoding.ASCII.GetBytes("input-prefix");
        using var input = new MemoryStream(inputPrefix.Concat(source).ToArray());
        input.Position = inputPrefix.Length;
        using var output = new MemoryStream();

        PdfPageExtractor.ExtractPages(input, output, 2);

        string text = NormalizeExtractedText(PdfReadDocument.Load(output.ToArray()).ExtractText());
        Assert.Contains("Secondpagemarker", text);
        Assert.DoesNotContain("Firstpagemarker", text);
    }

    [Fact]
    public void ExtractPageRange_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfPageExtractor.ExtractPageRange(source, output, 1, 2);

        byte[] extracted = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        string text = NormalizeExtractedText(PdfReadDocument.Load(extracted).ExtractText());
        Assert.Contains("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.DoesNotContain("Thirdpagemarker", text);
    }

    [Fact]
    public void ExtractPageRange_WritesStreamInputToOutputStream() {
        byte[] source = BuildThreePagePdf();
        byte[] inputPrefix = Encoding.ASCII.GetBytes("input-prefix");
        using var input = new MemoryStream(inputPrefix.Concat(source).ToArray());
        input.Position = inputPrefix.Length;
        using var output = new MemoryStream();

        PdfPageExtractor.ExtractPageRange(input, output, 2, 3);

        string text = NormalizeExtractedText(PdfReadDocument.Load(output.ToArray()).ExtractText());
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }

    [Fact]
    public void ExtractPageRange_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-range-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "range.pdf");
        string rangesOutputPath = Path.Combine(directory, "out", "ranges.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            PdfPageExtractor.ExtractPageRange(inputPath, outputPath, 1, 2);

            Assert.True(File.Exists(outputPath));
            PdfDocumentInfo info = PdfInspector.Inspect(outputPath);
            Assert.Equal(2, info.PageCount);

            string text = NormalizeExtractedText(PdfReadDocument.Load(outputPath).ExtractText());
            Assert.Contains("Firstpagemarker", text);
            Assert.Contains("Secondpagemarker", text);
            Assert.DoesNotContain("Thirdpagemarker", text);

            PdfPageExtractor.ExtractPageRanges(inputPath, rangesOutputPath, PdfPageRange.From(3, 3), PdfPageRange.From(1, 1));

            Assert.True(File.Exists(rangesOutputPath));
            PdfDocumentInfo rangesInfo = PdfInspector.Inspect(rangesOutputPath);
            Assert.Equal(2, rangesInfo.PageCount);

            string rangesText = NormalizeExtractedText(PdfReadDocument.Load(rangesOutputPath).ExtractText());
            AssertContainsInOrder(rangesText, "Thirdpagemarker", "Firstpagemarker");
            Assert.DoesNotContain("Secondpagemarker", rangesText);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractPathInputs_ReturnBytesForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-path-bytes-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            byte[] selected = PdfPageExtractor.ExtractPages(inputPath, 3, 1);
            string selectedText = NormalizeExtractedText(PdfReadDocument.Load(selected).ExtractText());
            Assert.Contains("Thirdpagemarker", selectedText);
            Assert.Contains("Firstpagemarker", selectedText);
            Assert.DoesNotContain("Secondpagemarker", selectedText);

            byte[] range = PdfPageExtractor.ExtractPageRange(inputPath, 2, 3);
            string rangeText = NormalizeExtractedText(PdfReadDocument.Load(range).ExtractText());
            Assert.DoesNotContain("Firstpagemarker", rangeText);
            Assert.Contains("Secondpagemarker", rangeText);
            Assert.Contains("Thirdpagemarker", rangeText);

            byte[] ranges = PdfPageExtractor.ExtractPageRanges(inputPath, PdfPageRange.From(3, 3), PdfPageRange.From(1, 1));
            string rangesText = NormalizeExtractedText(PdfReadDocument.Load(ranges).ExtractText());
            AssertContainsInOrder(rangesText, "Thirdpagemarker", "Firstpagemarker");
            Assert.DoesNotContain("Secondpagemarker", rangesText);

            IReadOnlyList<byte[]> split = PdfPageExtractor.SplitPages(inputPath);
            Assert.Equal(3, split.Count);
            for (int i = 0; i < split.Count; i++) {
                string splitText = NormalizeExtractedText(PdfReadDocument.Load(split[i]).ExtractText());
                Assert.Contains(NormalizeExtractedText(PageMarker(i + 1)), splitText);
            }

            IReadOnlyList<byte[]> rangeSplit = PdfPageExtractor.SplitPageRanges(inputPath, PdfPageRange.From(1, 2), PdfPageRange.From(3, 3));
            Assert.Equal(2, rangeSplit.Count);
            string rangeSplitText = NormalizeExtractedText(PdfReadDocument.Load(rangeSplit[0]).ExtractText());
            Assert.Contains("Firstpagemarker", rangeSplitText);
            Assert.Contains("Secondpagemarker", rangeSplitText);
            Assert.DoesNotContain("Thirdpagemarker", rangeSplitText);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractPathInputs_WriteToOutputStreamsForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-path-streams-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            using var selectedOutput = new MemoryStream();
            selectedOutput.Write(prefix, 0, prefix.Length);
            PdfPageExtractor.ExtractPages(inputPath, selectedOutput, 3, 1);
            byte[] selected = selectedOutput.ToArray().Skip(prefix.Length).ToArray();
            Assert.Equal(prefix, selectedOutput.ToArray().Take(prefix.Length).ToArray());
            string selectedText = NormalizeExtractedText(PdfReadDocument.Load(selected).ExtractText());
            AssertContainsInOrder(selectedText, "Thirdpagemarker", "Firstpagemarker");
            Assert.DoesNotContain("Secondpagemarker", selectedText);

            using var rangeOutput = new MemoryStream();
            rangeOutput.Write(prefix, 0, prefix.Length);
            PdfPageExtractor.ExtractPageRange(inputPath, rangeOutput, 2, 3);
            byte[] range = rangeOutput.ToArray().Skip(prefix.Length).ToArray();
            Assert.Equal(prefix, rangeOutput.ToArray().Take(prefix.Length).ToArray());
            string rangeText = NormalizeExtractedText(PdfReadDocument.Load(range).ExtractText());
            Assert.DoesNotContain("Firstpagemarker", rangeText);
            Assert.Contains("Secondpagemarker", rangeText);
            Assert.Contains("Thirdpagemarker", rangeText);

            using var rangeObjectOutput = new MemoryStream();
            rangeObjectOutput.Write(prefix, 0, prefix.Length);
            PdfPageExtractor.ExtractPageRange(inputPath, rangeObjectOutput, PdfPageRange.From(1, 2));
            byte[] rangeObject = rangeObjectOutput.ToArray().Skip(prefix.Length).ToArray();
            Assert.Equal(prefix, rangeObjectOutput.ToArray().Take(prefix.Length).ToArray());
            string rangeObjectText = NormalizeExtractedText(PdfReadDocument.Load(rangeObject).ExtractText());
            Assert.Contains("Firstpagemarker", rangeObjectText);
            Assert.Contains("Secondpagemarker", rangeObjectText);
            Assert.DoesNotContain("Thirdpagemarker", rangeObjectText);

            using var rangesOutput = new MemoryStream();
            rangesOutput.Write(prefix, 0, prefix.Length);
            PdfPageExtractor.ExtractPageRanges(inputPath, rangesOutput, PdfPageRange.From(3, 3), PdfPageRange.From(1, 1));
            byte[] ranges = rangesOutput.ToArray().Skip(prefix.Length).ToArray();
            Assert.Equal(prefix, rangesOutput.ToArray().Take(prefix.Length).ToArray());
            string rangesText = NormalizeExtractedText(PdfReadDocument.Load(ranges).ExtractText());
            AssertContainsInOrder(rangesText, "Thirdpagemarker", "Firstpagemarker");
            Assert.DoesNotContain("Secondpagemarker", rangesText);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void SplitPages_ReturnsSinglePageDocuments() {
        byte[] source = BuildThreePagePdf();

        IReadOnlyList<byte[]> pages = PdfPageExtractor.SplitPages(source);

        Assert.Equal(3, pages.Count);
        for (int i = 0; i < pages.Count; i++) {
            using var pdf = PdfDocument.Open(new MemoryStream(pages[i]));
            Assert.Equal(1, pdf.NumberOfPages);

            string text = NormalizeExtractedText(PdfReadDocument.Load(pages[i]).ExtractText());
            Assert.Contains(NormalizeExtractedText(PageMarker(i + 1)), text);
        }
    }

    [Fact]
    public void SplitPages_ReadsFromCurrentStreamPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        IReadOnlyList<byte[]> pages = PdfPageExtractor.SplitPages(stream);

        Assert.Equal(3, pages.Count);
        for (int i = 0; i < pages.Count; i++) {
            string text = NormalizeExtractedText(PdfReadDocument.Load(pages[i]).ExtractText());
            Assert.Contains(NormalizeExtractedText(PageMarker(i + 1)), text);
        }
    }

    [Fact]
    public void SplitPageRanges_ReturnsRangeDocuments() {
        byte[] source = BuildThreePagePdf();

        IReadOnlyList<byte[]> ranges = PdfPageExtractor.SplitPageRanges(
            source,
            PdfPageRange.From(1, 2),
            PdfPageRange.From(3, 3));

        Assert.Equal(2, ranges.Count);

        PdfDocumentInfo firstInfo = PdfInspector.Inspect(ranges[0]);
        Assert.Equal(2, firstInfo.PageCount);
        string firstText = NormalizeExtractedText(PdfReadDocument.Load(ranges[0]).ExtractText());
        Assert.Contains("Firstpagemarker", firstText);
        Assert.Contains("Secondpagemarker", firstText);
        Assert.DoesNotContain("Thirdpagemarker", firstText);

        PdfDocumentInfo secondInfo = PdfInspector.Inspect(ranges[1]);
        Assert.Equal(1, secondInfo.PageCount);
        string secondText = NormalizeExtractedText(PdfReadDocument.Load(ranges[1]).ExtractText());
        Assert.DoesNotContain("Firstpagemarker", secondText);
        Assert.DoesNotContain("Secondpagemarker", secondText);
        Assert.Contains("Thirdpagemarker", secondText);
    }

    [Fact]
    public void SplitPageRanges_ReadsFromCurrentStreamPosition() {
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        IReadOnlyList<byte[]> ranges = PdfPageExtractor.SplitPageRanges(stream, PdfPageRange.From(2, 3));

        Assert.Single(ranges);
        string text = NormalizeExtractedText(PdfReadDocument.Load(ranges[0]).ExtractText());
        Assert.DoesNotContain("Firstpagemarker", text);
        Assert.Contains("Secondpagemarker", text);
        Assert.Contains("Thirdpagemarker", text);
    }

    [Fact]
    public void SplitPages_WritesSinglePageDocumentsToDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-split-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputDirectory = Path.Combine(directory, "pages");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            IReadOnlyList<string> paths = PdfPageExtractor.SplitPages(inputPath, outputDirectory);

            Assert.Equal(3, paths.Count);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0001.pdf"), paths[0]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0002.pdf"), paths[1]);
            Assert.Equal(Path.Combine(outputDirectory, "source-page-0003.pdf"), paths[2]);

            for (int i = 0; i < paths.Count; i++) {
                Assert.True(File.Exists(paths[i]));
                PdfDocumentInfo info = PdfInspector.Inspect(paths[i]);
                Assert.Equal(1, info.PageCount);

                string text = NormalizeExtractedText(PdfReadDocument.Load(paths[i]).ExtractText());
                Assert.Contains(NormalizeExtractedText(PageMarker(i + 1)), text);
            }
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void SplitPageRanges_WritesRangeDocumentsToDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-split-ranges-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "source.pdf");
        string outputDirectory = Path.Combine(directory, "ranges");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            IReadOnlyList<string> paths = PdfPageExtractor.SplitPageRanges(
                inputPath,
                outputDirectory,
                PdfPageRange.From(1, 2),
                PdfPageRange.From(3, 3),
                PdfPageRange.From(1, 2));

            Assert.Equal(3, paths.Count);
            Assert.Equal(Path.Combine(outputDirectory, "source-pages-0001-0002.pdf"), paths[0]);
            Assert.Equal(Path.Combine(outputDirectory, "source-pages-0003-0003.pdf"), paths[1]);
            Assert.Equal(Path.Combine(outputDirectory, "source-pages-0001-0002-occurrence-0002.pdf"), paths[2]);
            Assert.NotEqual(paths[0], paths[2]);

            string firstText = NormalizeExtractedText(PdfReadDocument.Load(paths[0]).ExtractText());
            Assert.Contains("Firstpagemarker", firstText);
            Assert.Contains("Secondpagemarker", firstText);
            Assert.DoesNotContain("Thirdpagemarker", firstText);

            string secondText = NormalizeExtractedText(PdfReadDocument.Load(paths[1]).ExtractText());
            Assert.DoesNotContain("Firstpagemarker", secondText);
            Assert.DoesNotContain("Secondpagemarker", secondText);
            Assert.Contains("Thirdpagemarker", secondText);

            string duplicateText = NormalizeExtractedText(PdfReadDocument.Load(paths[2]).ExtractText());
            Assert.Contains("Firstpagemarker", duplicateText);
            Assert.Contains("Secondpagemarker", duplicateText);
            Assert.DoesNotContain("Thirdpagemarker", duplicateText);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void SplitPages_WritesStreamOutputWithDeterministicBaseName() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-stream-split-" + Guid.NewGuid().ToString("N"));
        string outputDirectory = Path.Combine(directory, "pages");
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        try {
            IReadOnlyList<string> paths = PdfPageExtractor.SplitPages(stream, outputDirectory, "stream-source.pdf");

            Assert.Equal(3, paths.Count);
            Assert.Equal(Path.Combine(outputDirectory, "stream-source-page-0001.pdf"), paths[0]);
            Assert.Equal(Path.Combine(outputDirectory, "stream-source-page-0002.pdf"), paths[1]);
            Assert.Equal(Path.Combine(outputDirectory, "stream-source-page-0003.pdf"), paths[2]);

            for (int i = 0; i < paths.Count; i++) {
                Assert.True(File.Exists(paths[i]));
                string text = NormalizeExtractedText(PdfReadDocument.Load(paths[i]).ExtractText());
                Assert.Contains(NormalizeExtractedText(PageMarker(i + 1)), text);
            }
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void SplitPageRanges_WritesStreamOutputWithDeterministicBaseName() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-stream-split-ranges-" + Guid.NewGuid().ToString("N"));
        string outputDirectory = Path.Combine(directory, "ranges");
        byte[] source = BuildThreePagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        try {
            IReadOnlyList<string> paths = PdfPageExtractor.SplitPageRanges(
                stream,
                outputDirectory,
                "stream-source.pdf",
                PdfPageRange.From(2, 3));

            Assert.Single(paths);
            Assert.Equal(Path.Combine(outputDirectory, "stream-source-pages-0002-0003.pdf"), paths[0]);

            string text = NormalizeExtractedText(PdfReadDocument.Load(paths[0]).ExtractText());
            Assert.DoesNotContain("Firstpagemarker", text);
            Assert.Contains("Secondpagemarker", text);
            Assert.Contains("Thirdpagemarker", text);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractPages_PreservesImageStreamsForSelectedPages() {
        byte[] source = PdfDoc.Create()
            .Paragraph(p => p.Text("Cover page"))
            .PageBreak()
            .Image(CreateMinimalRgbPng(), 24, 24)
            .Paragraph(p => p.Text("Image page marker"))
            .ToBytes();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 2);

        using var pdf = PdfDocument.Open(new MemoryStream(extracted));
        Assert.Equal(1, pdf.NumberOfPages);

        string pdfText = Encoding.ASCII.GetString(extracted);
        Assert.Contains("/Subtype /Image", pdfText);
        Assert.Contains("/Filter /FlateDecode", pdfText);
        Assert.Contains("/Length 12", pdfText);
        string extractedText = NormalizeExtractedText(PdfReadDocument.Load(extracted).ExtractText());
        Assert.Contains("Imagepagemarker", extractedText);
        Assert.DoesNotContain("Coverpage", extractedText);
    }

    [Fact]
    public void ExtractPages_PreservesLinkAnnotationsForSelectedPages() {
        byte[] source = PdfDoc.Create()
            .Paragraph(p => p.Text("Cover page"))
            .PageBreak()
            .Paragraph(p => p.Link("OfficeIMO link", "https://evotec.xyz"))
            .ToBytes();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.True(info.HasAnnotations);

        string pdfText = Encoding.ASCII.GetString(extracted);
        Assert.Contains("/Annots [", pdfText, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Link", pdfText, StringComparison.Ordinal);
        Assert.Contains("/URI (https://evotec.xyz)", pdfText, StringComparison.Ordinal);

        string extractedText = NormalizeExtractedText(PdfReadDocument.Load(extracted).ExtractText());
        Assert.Contains("OfficeIMOlink", extractedText);
        Assert.DoesNotContain("Coverpage", extractedText);
    }

    [Fact]
    public void ExtractPages_DropsBookmarkLinksWhenDestinationPageIsNotCopied() {
        byte[] source = PdfDoc.Create()
            .Paragraph(p => p.LinkToBookmark("Jump to details", "Details"))
            .PageBreak()
            .Bookmark("Details")
            .Paragraph(p => p.Text("Details marker"))
            .ToBytes();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 1);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Single(info.Pages);
        Assert.Empty(info.NamedDestinationNames);
        Assert.Empty(info.LinkAnnotations);
        Assert.Empty(info.Pages[0].LinkAnnotations);

        string pdfText = Encoding.ASCII.GetString(extracted);
        Assert.DoesNotContain("/S /GoTo", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("(Details)", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractPages_PreservesBookmarkLinksWhenDestinationPageIsCopied() {
        byte[] source = PdfDoc.Create()
            .Paragraph(p => p.LinkToBookmark("Jump to details", "Details"))
            .PageBreak()
            .Bookmark("Details")
            .Paragraph(p => p.Text("Details marker"))
            .ToBytes();

        byte[] extracted = PdfPageExtractor.ExtractPages(source, 1, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(extracted);
        Assert.Equal(2, info.PageCount);
        Assert.Equal(new[] { "Details" }, info.NamedDestinationNames);

        Assert.NotEmpty(info.LinkAnnotations);
        Assert.All(info.LinkAnnotations, link => {
            Assert.True(link.IsNamedDestinationLink);
            Assert.Equal("Details", link.DestinationName);
            Assert.Equal(1, link.PageNumber);
        });
    }

    [Fact]
    public void SplitPages_DropsBookmarkLinksWhoseDestinationsMoveToAnotherDocument() {
        byte[] source = PdfDoc.Create()
            .Paragraph(p => p.LinkToBookmark("Jump to details", "Details"))
            .PageBreak()
            .Bookmark("Details")
            .Paragraph(p => p.Text("Details marker"))
            .ToBytes();

        IReadOnlyList<byte[]> splitPages = PdfPageExtractor.SplitPages(source);

        Assert.Equal(2, splitPages.Count);
        PdfDocumentInfo first = PdfInspector.Inspect(splitPages[0]);
        PdfDocumentInfo second = PdfInspector.Inspect(splitPages[1]);

        Assert.Empty(first.LinkAnnotations);
        Assert.Empty(first.NamedDestinationNames);
        Assert.Empty(second.LinkAnnotations);
        Assert.Equal(new[] { "Details" }, second.NamedDestinationNames);
    }

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
        Assert.Throws<ArgumentNullException>(() => PdfPageExtractor.ExtractPageRanges(source, null!, PdfPageRange.From(1, 1)));
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

    private static byte[] BuildThreePagePdf() {
        var doc = PdfDoc.Create()
            .Meta(
                title: "Extraction sample",
                author: "OfficeIMO",
                subject: "Manipulation",
                keywords: "pdf,extract,split");

        doc.Compose(compose => {
            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(PageMarker(1)))));
            });

            compose.Page(page => {
                page.Size(new PageSize(612, 792));
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(PageMarker(2)))));
            });

            compose.Page(page => {
                page.Size(new PageSize(300, 500));
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(PageMarker(3)))));
            });
        });

        return doc.ToBytes();
    }

    private static string PageMarker(int pageNumber) {
        return pageNumber switch {
            1 => "First page marker",
            2 => "Second page marker",
            3 => "Third page marker",
            _ => "Page " + pageNumber
        };
    }

    private static string NormalizeExtractedText(string text) {
        return text.Replace(" ", string.Empty);
    }

    private static void AssertContainsInOrder(string text, params string[] expected) {
        int previous = -1;
        foreach (string item in expected) {
            int index = text.IndexOf(item, previous + 1, StringComparison.Ordinal);
            Assert.True(index >= 0, "Expected text '" + item + "' was not found after index " + previous + " in '" + text + "'.");
            previous = index;
        }
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }

    private sealed class WriteOnlyStream : MemoryStream {
        public override bool CanRead => false;
    }

    private sealed class ReadOnlyStream : MemoryStream {
        public override bool CanWrite => false;
    }
}

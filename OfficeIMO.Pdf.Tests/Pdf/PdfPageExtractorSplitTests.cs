using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageExtractorTests {
    [Fact]
    public void SplitPages_ReturnsSinglePageDocuments() {
        byte[] source = BuildThreePagePdf();

        IReadOnlyList<byte[]> pages = PdfPageExtractor.SplitPages(source);

        Assert.Equal(3, pages.Count);
        for (int i = 0; i < pages.Count; i++) {
            using var pdf = PdfPigDocument.Open(new MemoryStream(pages[i]));
            Assert.Equal(1, pdf.NumberOfPages);

            string text = NormalizeExtractedText(PdfReadDocument.Open(pages[i]).ExtractText());
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
            string text = NormalizeExtractedText(PdfReadDocument.Open(pages[i]).ExtractText());
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
        string firstText = NormalizeExtractedText(PdfReadDocument.Open(ranges[0]).ExtractText());
        Assert.Contains("Firstpagemarker", firstText);
        Assert.Contains("Secondpagemarker", firstText);
        Assert.DoesNotContain("Thirdpagemarker", firstText);

        PdfDocumentInfo secondInfo = PdfInspector.Inspect(ranges[1]);
        Assert.Equal(1, secondInfo.PageCount);
        string secondText = NormalizeExtractedText(PdfReadDocument.Open(ranges[1]).ExtractText());
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
        string text = NormalizeExtractedText(PdfReadDocument.Open(ranges[0]).ExtractText());
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

                string text = NormalizeExtractedText(PdfReadDocument.Open(paths[i]).ExtractText());
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

            string firstText = NormalizeExtractedText(PdfReadDocument.Open(paths[0]).ExtractText());
            Assert.Contains("Firstpagemarker", firstText);
            Assert.Contains("Secondpagemarker", firstText);
            Assert.DoesNotContain("Thirdpagemarker", firstText);

            string secondText = NormalizeExtractedText(PdfReadDocument.Open(paths[1]).ExtractText());
            Assert.DoesNotContain("Firstpagemarker", secondText);
            Assert.DoesNotContain("Secondpagemarker", secondText);
            Assert.Contains("Thirdpagemarker", secondText);

            string duplicateText = NormalizeExtractedText(PdfReadDocument.Open(paths[2]).ExtractText());
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
                string text = NormalizeExtractedText(PdfReadDocument.Open(paths[i]).ExtractText());
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

            string text = NormalizeExtractedText(PdfReadDocument.Open(paths[0]).ExtractText());
            Assert.DoesNotContain("Firstpagemarker", text);
            Assert.Contains("Secondpagemarker", text);
            Assert.Contains("Thirdpagemarker", text);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}

using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageExtractorTests {
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

            string text = NormalizeExtractedText(PdfReadDocument.Open(outputPath).ExtractText());
            Assert.Contains("Firstpagemarker", text);
            Assert.Contains("Secondpagemarker", text);
            Assert.DoesNotContain("Thirdpagemarker", text);

            PdfPageExtractor.ExtractPageRanges(inputPath, rangesOutputPath, PdfPageRange.From(3, 3), PdfPageRange.From(1, 1));

            Assert.True(File.Exists(rangesOutputPath));
            PdfDocumentInfo rangesInfo = PdfInspector.Inspect(rangesOutputPath);
            Assert.Equal(2, rangesInfo.PageCount);

            string rangesText = NormalizeExtractedText(PdfReadDocument.Open(rangesOutputPath).ExtractText());
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
            string selectedText = NormalizeExtractedText(PdfReadDocument.Open(selected).ExtractText());
            Assert.Contains("Thirdpagemarker", selectedText);
            Assert.Contains("Firstpagemarker", selectedText);
            Assert.DoesNotContain("Secondpagemarker", selectedText);

            byte[] range = PdfPageExtractor.ExtractPageRange(inputPath, 2, 3);
            string rangeText = NormalizeExtractedText(PdfReadDocument.Open(range).ExtractText());
            Assert.DoesNotContain("Firstpagemarker", rangeText);
            Assert.Contains("Secondpagemarker", rangeText);
            Assert.Contains("Thirdpagemarker", rangeText);

            byte[] ranges = PdfPageExtractor.ExtractPageRanges(inputPath, PdfPageRange.From(3, 3), PdfPageRange.From(1, 1));
            string rangesText = NormalizeExtractedText(PdfReadDocument.Open(ranges).ExtractText());
            AssertContainsInOrder(rangesText, "Thirdpagemarker", "Firstpagemarker");
            Assert.DoesNotContain("Secondpagemarker", rangesText);

            IReadOnlyList<byte[]> split = PdfPageExtractor.SplitPages(inputPath);
            Assert.Equal(3, split.Count);
            for (int i = 0; i < split.Count; i++) {
                string splitText = NormalizeExtractedText(PdfReadDocument.Open(split[i]).ExtractText());
                Assert.Contains(NormalizeExtractedText(PageMarker(i + 1)), splitText);
            }

            IReadOnlyList<byte[]> rangeSplit = PdfPageExtractor.SplitPageRanges(inputPath, PdfPageRange.From(1, 2), PdfPageRange.From(3, 3));
            Assert.Equal(2, rangeSplit.Count);
            string rangeSplitText = NormalizeExtractedText(PdfReadDocument.Open(rangeSplit[0]).ExtractText());
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
            string selectedText = NormalizeExtractedText(PdfReadDocument.Open(selected).ExtractText());
            AssertContainsInOrder(selectedText, "Thirdpagemarker", "Firstpagemarker");
            Assert.DoesNotContain("Secondpagemarker", selectedText);

            using var rangeOutput = new MemoryStream();
            rangeOutput.Write(prefix, 0, prefix.Length);
            PdfPageExtractor.ExtractPageRange(inputPath, rangeOutput, 2, 3);
            byte[] range = rangeOutput.ToArray().Skip(prefix.Length).ToArray();
            Assert.Equal(prefix, rangeOutput.ToArray().Take(prefix.Length).ToArray());
            string rangeText = NormalizeExtractedText(PdfReadDocument.Open(range).ExtractText());
            Assert.DoesNotContain("Firstpagemarker", rangeText);
            Assert.Contains("Secondpagemarker", rangeText);
            Assert.Contains("Thirdpagemarker", rangeText);

            using var rangeObjectOutput = new MemoryStream();
            rangeObjectOutput.Write(prefix, 0, prefix.Length);
            PdfPageExtractor.ExtractPageRange(inputPath, rangeObjectOutput, PdfPageRange.From(1, 2));
            byte[] rangeObject = rangeObjectOutput.ToArray().Skip(prefix.Length).ToArray();
            Assert.Equal(prefix, rangeObjectOutput.ToArray().Take(prefix.Length).ToArray());
            string rangeObjectText = NormalizeExtractedText(PdfReadDocument.Open(rangeObject).ExtractText());
            Assert.Contains("Firstpagemarker", rangeObjectText);
            Assert.Contains("Secondpagemarker", rangeObjectText);
            Assert.DoesNotContain("Thirdpagemarker", rangeObjectText);

            using var rangesOutput = new MemoryStream();
            rangesOutput.Write(prefix, 0, prefix.Length);
            PdfPageExtractor.ExtractPageRanges(inputPath, rangesOutput, PdfPageRange.From(3, 3), PdfPageRange.From(1, 1));
            byte[] ranges = rangesOutput.ToArray().Skip(prefix.Length).ToArray();
            Assert.Equal(prefix, rangesOutput.ToArray().Take(prefix.Length).ToArray());
            string rangesText = NormalizeExtractedText(PdfReadDocument.Open(ranges).ExtractText());
            AssertContainsInOrder(rangesText, "Thirdpagemarker", "Firstpagemarker");
            Assert.DoesNotContain("Secondpagemarker", rangesText);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}

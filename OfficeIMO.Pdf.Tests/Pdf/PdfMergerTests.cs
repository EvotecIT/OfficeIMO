using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfMergerTests {
    [Fact]
    public void Merge_CombinesPagesInSourceOrder() {
        byte[] first = BuildPdf(
            "First document",
            "First page",
            ("First second page", new PageSize(300, 500)));
        byte[] second = BuildPdf(
            "Second document",
            "Second document page",
            ("Second second page", new PageSize(792, 612)));

        byte[] merged = PdfMerger.Merge(first, second);

        using var pdf = PdfPigDocument.Open(new MemoryStream(merged));
        Assert.Equal(4, pdf.NumberOfPages);

        var read = PdfReadDocument.Open(merged);
        string text = NormalizeExtractedText(read.ExtractText());
        AssertContainsInOrder(text, "Firstpage", "Firstsecondpage", "Seconddocumentpage", "Secondsecondpage");

        PdfDocumentInfo info = PdfInspector.Inspect(merged);
        Assert.Equal(4, info.PageCount);
        Assert.Equal("First document", info.Metadata.Title);
        Assert.Equal("OfficeIMO", info.Metadata.Author);
        Assert.Equal(612, info.Pages[0].Width);
        Assert.Equal(792, info.Pages[0].Height);
        Assert.Equal(300, info.Pages[1].Width);
        Assert.Equal(500, info.Pages[1].Height);
        Assert.Equal(612, info.Pages[2].Width);
        Assert.Equal(792, info.Pages[2].Height);
        Assert.Equal(792, info.Pages[3].Width);
        Assert.Equal(612, info.Pages[3].Height);
    }

    [Fact]
    public void Merge_PreservesImageStreams() {
        byte[] withImage = PdfDocument.Create()
            .Meta(title: "Image source", author: "OfficeIMO")
            .Image(CreateMinimalRgbPng(), 24, 24)
            .Paragraph(p => p.Text("Image source page"))
            .ToBytes();
        byte[] textOnly = BuildPdf("Text source", "Text only page");

        byte[] merged = PdfMerger.Merge(withImage, textOnly);

        using var pdf = PdfPigDocument.Open(new MemoryStream(merged));
        Assert.Equal(2, pdf.NumberOfPages);

        string pdfText = Encoding.ASCII.GetString(merged);
        Assert.Contains("/Subtype /Image", pdfText);
        Assert.Contains("/Filter /FlateDecode", pdfText);
        Assert.Contains("/Width 1", pdfText);
        Assert.Contains("/Height 1", pdfText);

        string text = NormalizeExtractedText(PdfReadDocument.Open(merged).ExtractText());
        AssertContainsInOrder(text, "Imagesourcepage", "Textonlypage");
    }

    [Fact]
    public void Merge_WithFlattenVisualAnnotationsOption_PaintsVisualAnnotationsAndRemovesLiveAnnotations() {
        byte[] annotated = BuildAnnotatedPdf("Annotated source", "Annotated merge source");
        byte[] plain = BuildPdf("Plain source", "Plain merge page");

        byte[] merged = PdfMerger.Merge(
            new PdfMergeOptions {
                FlattenVisualAnnotations = true
            },
            annotated,
            plain);
        string pdf = Encoding.ASCII.GetString(merged);
        PdfDocumentInfo info = PdfInspector.Inspect(merged);

        Assert.Equal(2, info.PageCount);
        Assert.False(info.HasAnnotations);
        Assert.Equal(0, info.AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Highlight", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot2 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("BT\n/Helv", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0.9 0.1 rg 0 0 120 14 re f", pdf, StringComparison.Ordinal);

        string text = NormalizeExtractedText(PdfReadDocument.Open(merged).ExtractText());
        AssertContainsInOrder(text, "Annotatedmergesource", "Plainmergepage");
    }

    [Fact]
    public void Merge_WithResizePagesOption_NormalizesMixedSourcePageSizes() {
        byte[] first = BuildPdf(
            "Mixed size first",
            "Letter source page",
            ("Small source page", new PageSize(300, 500)));
        byte[] second = BuildPdf(
            "Mixed size second",
            "Landscape source page",
            ("A4 source page", PageSizes.A4));

        byte[] merged = PdfMerger.Merge(
            new PdfMergeOptions {
                ResizePages = new PdfPageResizeOptions(PageSizes.A4) {
                    Margin = 12,
                    Mode = PdfPageResizeMode.Fit
                }
            },
            first,
            second);

        PdfDocumentInfo info = PdfInspector.Inspect(merged);
        Assert.Equal(4, info.PageCount);
        Assert.All(info.Pages, page => {
            Assert.Equal(595, Math.Round(page.Width));
            Assert.Equal(842, Math.Round(page.Height));
        });

        string pdf = Encoding.ASCII.GetString(merged);
        Assert.Contains(" cm", pdf, StringComparison.Ordinal);
        string text = NormalizeExtractedText(PdfReadDocument.Open(merged).ExtractText());
        AssertContainsInOrder(text, "Lettersourcepage", "Smallsourcepage", "Landscapesourcepage", "A4sourcepage");
    }

    [Fact]
    public void Merge_ReadsStreamsFromCurrentPositions() {
        using var first = CreatePrefixedStream(BuildPdf("First stream", "First stream page"));
        using var second = CreatePrefixedStream(BuildPdf("Second stream", "Second stream page"));

        byte[] merged = PdfMerger.Merge(first, second);

        PdfDocumentInfo info = PdfInspector.Inspect(merged);
        Assert.Equal(2, info.PageCount);
        Assert.Equal("First stream", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(merged).ExtractText());
        AssertContainsInOrder(text, "Firststreampage", "Secondstreampage");
    }

    [Fact]
    public void Merge_ReadsEnumerableStreamsFromCurrentPositions() {
        using var first = CreatePrefixedStream(BuildPdf("First enumerable stream", "First enumerable page"));
        using var second = CreatePrefixedStream(BuildPdf("Second enumerable stream", "Second enumerable page"));

        byte[] merged = PdfMerger.Merge((IEnumerable<Stream>)new[] { first, second });

        PdfDocumentInfo info = PdfInspector.Inspect(merged);
        Assert.Equal(2, info.PageCount);

        string text = NormalizeExtractedText(PdfReadDocument.Open(merged).ExtractText());
        AssertContainsInOrder(text, "Firstenumerablepage", "Secondenumerablepage");
    }

    [Fact]
    public void Merge_WritesByteInputsToOutputStreamAtCurrentPosition() {
        byte[] first = BuildPdf("First output", "First output page");
        byte[] second = BuildPdf("Second output", "Second output page");
        using var output = CreateOutputStream(out int prefixLength);

        PdfMerger.Merge((IEnumerable<byte[]>)new[] { first, second }, output);

        byte[] merged = GetOutputPayload(output, prefixLength);
        PdfDocumentInfo info = PdfInspector.Inspect(merged);
        Assert.Equal(2, info.PageCount);
        Assert.Equal("First output", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(merged).ExtractText());
        AssertContainsInOrder(text, "Firstoutputpage", "Secondoutputpage");
    }

    [Fact]
    public void Merge_WritesStreamInputsToOutputStreamAtCurrentPosition() {
        using var first = CreatePrefixedStream(BuildPdf("First stream output", "First stream output page"));
        using var second = CreatePrefixedStream(BuildPdf("Second stream output", "Second stream output page"));
        using var output = CreateOutputStream(out int prefixLength);

        PdfMerger.Merge((IEnumerable<Stream>)new[] { first, second }, output);

        byte[] merged = GetOutputPayload(output, prefixLength);
        PdfDocumentInfo info = PdfInspector.Inspect(merged);
        Assert.Equal(2, info.PageCount);
        Assert.Equal("First stream output", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(merged).ExtractText());
        AssertContainsInOrder(text, "Firststreamoutputpage", "Secondstreamoutputpage");
    }

    [Fact]
    public void MergeFiles_WritesMergedPdf() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-merge-" + Guid.NewGuid().ToString("N"));
        string firstPath = Path.Combine(directory, "first.pdf");
        string secondPath = Path.Combine(directory, "second.pdf");
        string outputPath = Path.Combine(directory, "merged", "merged.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(firstPath, BuildPdf("First", "File first"));
            File.WriteAllBytes(secondPath, BuildPdf("Second", "File second"));

            PdfMerger.MergeFiles(outputPath, firstPath, secondPath);

            Assert.True(File.Exists(outputPath));
            var info = PdfInspector.Inspect(outputPath);
            Assert.Equal(2, info.PageCount);
            string text = NormalizeExtractedText(PdfReadDocument.Open(outputPath).ExtractText());
            AssertContainsInOrder(text, "Filefirst", "Filesecond");
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void MergeFiles_WritesEnumerableInputPathsToOutputPath() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-merge-enumerable-" + Guid.NewGuid().ToString("N"));
        string firstPath = Path.Combine(directory, "first.pdf");
        string secondPath = Path.Combine(directory, "second.pdf");
        string outputPath = Path.Combine(directory, "merged", "merged-enumerable.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(firstPath, BuildPdf("First enumerable path", "File first enumerable"));
            File.WriteAllBytes(secondPath, BuildPdf("Second enumerable path", "File second enumerable"));

            PdfMerger.MergeFiles((IEnumerable<string>)new[] { firstPath, secondPath }, outputPath);

            Assert.True(File.Exists(outputPath));
            var info = PdfInspector.Inspect(outputPath);
            Assert.Equal(2, info.PageCount);
            Assert.Equal("First enumerable path", info.Metadata.Title);
            string text = NormalizeExtractedText(PdfReadDocument.Open(outputPath).ExtractText());
            AssertContainsInOrder(text, "Filefirstenumerable", "Filesecondenumerable");
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void MergeFiles_WritesEnumerableInputPathsToOutputStreamAtCurrentPosition() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-merge-enumerable-stream-" + Guid.NewGuid().ToString("N"));
        string firstPath = Path.Combine(directory, "first.pdf");
        string secondPath = Path.Combine(directory, "second.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(firstPath, BuildPdf("First file stream", "File first stream"));
            File.WriteAllBytes(secondPath, BuildPdf("Second file stream", "File second stream"));
            using var output = CreateOutputStream(out int prefixLength);

            PdfMerger.MergeFiles((IEnumerable<string>)new[] { firstPath, secondPath }, output);

            byte[] merged = GetOutputPayload(output, prefixLength);
            PdfDocumentInfo info = PdfInspector.Inspect(merged);
            Assert.Equal(2, info.PageCount);
            Assert.Equal("First file stream", info.Metadata.Title);
            string text = NormalizeExtractedText(PdfReadDocument.Open(merged).ExtractText());
            AssertContainsInOrder(text, "Filefirststream", "Filesecondstream");
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void MergeFilesToBytes_ReturnsMergedPdfForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-merge-bytes-" + Guid.NewGuid().ToString("N"));
        string firstPath = Path.Combine(directory, "first.pdf");
        string secondPath = Path.Combine(directory, "second.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(firstPath, BuildPdf("First bytes", "File first bytes"));
            File.WriteAllBytes(secondPath, BuildPdf("Second bytes", "File second bytes"));

            byte[] mergedFromParams = PdfMerger.MergeFilesToBytes(firstPath, secondPath);
            PdfDocumentInfo paramsInfo = PdfInspector.Inspect(mergedFromParams);
            Assert.Equal(2, paramsInfo.PageCount);
            Assert.Equal("First bytes", paramsInfo.Metadata.Title);
            string paramsText = NormalizeExtractedText(PdfReadDocument.Open(mergedFromParams).ExtractText());
            AssertContainsInOrder(paramsText, "Filefirstbytes", "Filesecondbytes");

            byte[] mergedFromEnumerable = PdfMerger.MergeFilesToBytes((IEnumerable<string>)new[] { firstPath, secondPath });
            PdfDocumentInfo enumerableInfo = PdfInspector.Inspect(mergedFromEnumerable);
            Assert.Equal(2, enumerableInfo.PageCount);
            string enumerableText = NormalizeExtractedText(PdfReadDocument.Open(mergedFromEnumerable).ExtractText());
            AssertContainsInOrder(enumerableText, "Filefirstbytes", "Filesecondbytes");
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void Merge_RejectsInvalidInputs() {
        Assert.Throws<ArgumentException>(() => PdfMerger.Merge(Array.Empty<byte[]>()));
        Assert.Throws<ArgumentException>(() => PdfMerger.Merge(new byte[][] { null! }));
        Assert.Throws<ArgumentNullException>(() => PdfMerger.Merge((IEnumerable<byte[]>)null!));
        Assert.Throws<ArgumentException>(() => PdfMerger.Merge(Array.Empty<Stream>()));
        Assert.Throws<ArgumentException>(() => PdfMerger.Merge(new Stream[] { null! }));
        Assert.Throws<ArgumentException>(() => PdfMerger.Merge(new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfMerger.Merge((IEnumerable<Stream>)null!));
        byte[] source = BuildPdf("Source", "Source page");
        Assert.Throws<ArgumentNullException>(() => PdfMerger.Merge((IEnumerable<byte[]>)new[] { source }, null!));
        Assert.Throws<ArgumentException>(() => PdfMerger.Merge((IEnumerable<byte[]>)new[] { source }, new ReadOnlyStream()));
        using var nullOutputInput = new MemoryStream(source);
        using var readOnlyOutputInput = new MemoryStream(source);
        Assert.Throws<ArgumentNullException>(() => PdfMerger.Merge((IEnumerable<Stream>)new[] { nullOutputInput }, null!));
        Assert.Throws<ArgumentException>(() => PdfMerger.Merge((IEnumerable<Stream>)new[] { readOnlyOutputInput }, new ReadOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfMerger.MergeFiles((string)null!, "input.pdf"));
        Assert.Throws<ArgumentException>(() => PdfMerger.MergeFiles(" ", "input.pdf"));
        Assert.Throws<ArgumentNullException>(() => PdfMerger.MergeFiles("output.pdf", null!));
        Assert.Throws<ArgumentException>(() => PdfMerger.MergeFiles("output.pdf", " "));
        Assert.Throws<ArgumentNullException>(() => PdfMerger.MergeFiles((IEnumerable<string>)null!, "output.pdf"));
        Assert.Throws<ArgumentException>(() => PdfMerger.MergeFiles(Array.Empty<string>(), "output.pdf"));
        Assert.Throws<ArgumentException>(() => PdfMerger.MergeFiles(new string[] { null! }, "output.pdf"));
        Assert.Throws<ArgumentException>(() => PdfMerger.MergeFiles(new[] { " " }, "output.pdf"));
        Assert.Throws<ArgumentNullException>(() => PdfMerger.MergeFiles(new[] { "input.pdf" }, (string)null!));
        Assert.Throws<ArgumentException>(() => PdfMerger.MergeFiles(new[] { "input.pdf" }, " "));
        Assert.Throws<ArgumentNullException>(() => PdfMerger.MergeFiles((IEnumerable<string>)new[] { "input.pdf" }, (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfMerger.MergeFiles((IEnumerable<string>)new[] { "input.pdf" }, new ReadOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfMerger.MergeFilesToBytes((string[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfMerger.MergeFilesToBytes((IEnumerable<string>)null!));
        Assert.Throws<ArgumentException>(() => PdfMerger.MergeFilesToBytes(Array.Empty<string>()));
        Assert.Throws<ArgumentException>(() => PdfMerger.MergeFilesToBytes(new string[] { null! }));
        Assert.Throws<ArgumentException>(() => PdfMerger.MergeFilesToBytes(" "));
    }

    [Fact]
    public void MergeFiles_RejectsDirectoryOutputTargets() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-merge-output-path-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputDirectory = Path.Combine(directory, "existing-output");

        try {
            Directory.CreateDirectory(outputDirectory);
            File.WriteAllBytes(inputPath, BuildPdf("Input", "Input page"));

            var exception = Assert.Throws<ArgumentException>(() =>
                PdfMerger.MergeFiles(outputDirectory, inputPath));

            Assert.Equal("outputPath", exception.ParamName);
            Assert.Contains("Output path refers to a directory; a file path is required.", exception.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static byte[] BuildPdf(string title, string firstPageText, params (string Text, PageSize Size)[] extraPages) {
        var doc = PdfDocument.Create()
            .Meta(title: title, author: "OfficeIMO")
            .Paragraph(p => p.Text(firstPageText));

        if (extraPages.Length > 0) {
            doc.Compose(compose => {
                foreach (var extraPage in extraPages) {
                    compose.Page(page => {
                        page.Size(extraPage.Size);
                        page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(extraPage.Text))));
                    });
                }
            });
        }

        return doc.ToBytes();
    }

    private static byte[] BuildAnnotatedPdf(string title, string pageText) {
        return PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Meta(title: title, author: "OfficeIMO")
            .Paragraph(p => p.Text(pageText))
            .FreeTextAnnotation(
                "Merge review note",
                width: 150,
                height: 44,
                borderColor: new PdfColor(0.2D, 0.4D, 0.8D),
                fillColor: new PdfColor(0.95D, 0.98D, 1D),
                textAlign: PdfAlign.Center)
            .HighlightAnnotation("Merge highlight", width: 120, height: 14, color: new PdfColor(1D, 0.9D, 0.1D))
            .ToBytes();
    }

    private static void AssertContainsInOrder(string text, params string[] expected) {
        int previous = -1;
        foreach (string item in expected) {
            int index = text.IndexOf(item, StringComparison.Ordinal);
            Assert.True(index >= 0, "Expected text '" + item + "' was not found in '" + text + "'.");
            Assert.True(index > previous, "Expected text '" + item + "' to appear after the previous marker.");
            previous = index;
        }
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

    private static byte[] CreateMinimalRgbPng() => PdfPngTestImages.CreateRgbPng(255, 0, 0);
}

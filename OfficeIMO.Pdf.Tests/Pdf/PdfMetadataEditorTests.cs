using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfMetadataEditorTests {
    [Fact]
    public void UpdateMetadata_PreservesUnspecifiedFieldsAndPages() {
        byte[] source = BuildTwoPagePdf();

        byte[] edited = PdfMetadataEditor.UpdateMetadata(
            source,
            title: "Updated title",
            keywords: "updated,metadata");

        using var pdf = PdfPigDocument.Open(new MemoryStream(edited));
        Assert.Equal(2, pdf.NumberOfPages);

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal(2, info.PageCount);
        Assert.Equal("Updated title", info.Metadata.Title);
        Assert.Equal("Original author", info.Metadata.Author);
        Assert.Equal("Original subject", info.Metadata.Subject);
        Assert.Equal("updated,metadata", info.Metadata.Keywords);
        Assert.Equal(595, info.Pages[0].Width);
        Assert.Equal(842, info.Pages[0].Height);
        Assert.Equal(300, info.Pages[1].Width);
        Assert.Equal(500, info.Pages[1].Height);

        string text = NormalizeExtractedText(PdfReadDocument.Open(edited).ExtractText());
        Assert.Contains("Firstpage", text);
        Assert.Contains("Secondpage", text);
    }

    [Fact]
    public void UpdateMetadata_EmptyStringClearsField() {
        byte[] source = BuildTwoPagePdf();

        byte[] edited = PdfMetadataEditor.UpdateMetadata(source, author: string.Empty);

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Original title", info.Metadata.Title);
        Assert.Null(info.Metadata.Author);
        Assert.Equal("Original subject", info.Metadata.Subject);
    }

    [Fact]
    public void UpdateMetadata_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildTwoPagePdf());

        byte[] edited = PdfMetadataEditor.UpdateMetadata(
            stream,
            title: "Stream title",
            keywords: "stream,metadata");

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Stream title", info.Metadata.Title);
        Assert.Equal("Original author", info.Metadata.Author);
        Assert.Equal("Original subject", info.Metadata.Subject);
        Assert.Equal("stream,metadata", info.Metadata.Keywords);
        Assert.Equal(2, info.PageCount);
    }

    [Fact]
    public void UpdateMetadata_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildTwoPagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfMetadataEditor.UpdateMetadata(source, output, title: "Output title");

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Output title", info.Metadata.Title);
        Assert.Equal("Original author", info.Metadata.Author);
        Assert.Equal(2, info.PageCount);
    }

    [Fact]
    public void UpdateMetadata_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildTwoPagePdf());
        using var output = new MemoryStream();

        PdfMetadataEditor.UpdateMetadata(input, output, author: "Stream output author");

        PdfDocumentInfo info = PdfInspector.Inspect(output.ToArray());
        Assert.Equal("Original title", info.Metadata.Title);
        Assert.Equal("Stream output author", info.Metadata.Author);
        Assert.Equal(2, info.PageCount);
    }

    [Fact]
    public void ReplaceMetadata_ReplacesAllFields() {
        byte[] source = BuildTwoPagePdf();

        byte[] edited = PdfMetadataEditor.ReplaceMetadata(source, new PdfMetadata {
            Title = "Replacement title",
            Author = "Replacement author"
        });

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Replacement title", info.Metadata.Title);
        Assert.Equal("Replacement author", info.Metadata.Author);
        Assert.Null(info.Metadata.Subject);
        Assert.Null(info.Metadata.Keywords);
    }

    [Fact]
    public void ReplaceMetadata_ReadsFromCurrentStreamPosition() {
        using var stream = CreatePrefixedStream(BuildTwoPagePdf());

        byte[] edited = PdfMetadataEditor.ReplaceMetadata(stream, new PdfMetadata {
            Title = "Stream replacement",
            Subject = "Stream subject"
        });

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Stream replacement", info.Metadata.Title);
        Assert.Null(info.Metadata.Author);
        Assert.Equal("Stream subject", info.Metadata.Subject);
        Assert.Null(info.Metadata.Keywords);
        Assert.Equal(2, info.PageCount);

        string text = NormalizeExtractedText(PdfReadDocument.Open(edited).ExtractText());
        Assert.Contains("Firstpage", text);
        Assert.Contains("Secondpage", text);
    }

    [Fact]
    public void ReplaceMetadata_WritesToOutputStreamAtCurrentPosition() {
        byte[] source = BuildTwoPagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var output = new MemoryStream();
        output.Write(prefix, 0, prefix.Length);

        PdfMetadataEditor.ReplaceMetadata(source, output, new PdfMetadata {
            Title = "Output replacement",
            Keywords = "output,metadata"
        });

        byte[] edited = output.ToArray().Skip(prefix.Length).ToArray();
        Assert.Equal(prefix, output.ToArray().Take(prefix.Length).ToArray());

        PdfDocumentInfo info = PdfInspector.Inspect(edited);
        Assert.Equal("Output replacement", info.Metadata.Title);
        Assert.Null(info.Metadata.Author);
        Assert.Null(info.Metadata.Subject);
        Assert.Equal("output,metadata", info.Metadata.Keywords);
    }

    [Fact]
    public void ReplaceMetadata_WritesStreamInputToOutputStream() {
        using var input = CreatePrefixedStream(BuildTwoPagePdf());
        using var output = new MemoryStream();

        PdfMetadataEditor.ReplaceMetadata(input, output, new PdfMetadata {
            Author = "Output stream author",
            Subject = "Output stream subject"
        });

        PdfDocumentInfo info = PdfInspector.Inspect(output.ToArray());
        Assert.Null(info.Metadata.Title);
        Assert.Equal("Output stream author", info.Metadata.Author);
        Assert.Equal("Output stream subject", info.Metadata.Subject);
        Assert.Null(info.Metadata.Keywords);
    }

    [Fact]
    public void UpdateMetadata_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-metadata-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "metadata.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildTwoPagePdf());

            PdfMetadataEditor.UpdateMetadata(inputPath, outputPath, title: "Path title");

            Assert.True(File.Exists(outputPath));
            PdfDocumentInfo info = PdfInspector.Inspect(outputPath);
            Assert.Equal("Path title", info.Metadata.Title);
            Assert.Equal("Original author", info.Metadata.Author);
            Assert.Equal(2, info.PageCount);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ReplaceMetadata_WritesPathOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-metadata-replace-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputPath = Path.Combine(directory, "out", "metadata.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildTwoPagePdf());

            PdfMetadataEditor.ReplaceMetadata(inputPath, outputPath, new PdfMetadata {
                Title = "Path replacement",
                Author = "Path author"
            });

            Assert.True(File.Exists(outputPath));
            PdfDocumentInfo info = PdfInspector.Inspect(outputPath);
            Assert.Equal("Path replacement", info.Metadata.Title);
            Assert.Equal("Path author", info.Metadata.Author);
            Assert.Null(info.Metadata.Subject);
            Assert.Equal(2, info.PageCount);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void MetadataPathInputs_ReturnBytesForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-metadata-path-bytes-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildTwoPagePdf());

            byte[] updated = PdfMetadataEditor.UpdateMetadataToBytes(
                inputPath,
                title: "Path bytes title",
                keywords: "path,bytes");
            PdfDocumentInfo updatedInfo = PdfInspector.Inspect(updated);
            Assert.Equal("Path bytes title", updatedInfo.Metadata.Title);
            Assert.Equal("Original author", updatedInfo.Metadata.Author);
            Assert.Equal("Original subject", updatedInfo.Metadata.Subject);
            Assert.Equal("path,bytes", updatedInfo.Metadata.Keywords);
            Assert.Equal(2, updatedInfo.PageCount);

            byte[] replaced = PdfMetadataEditor.ReplaceMetadataToBytes(inputPath, new PdfMetadata {
                Title = "Path bytes replacement",
                Author = "Path bytes author"
            });
            PdfDocumentInfo replacedInfo = PdfInspector.Inspect(replaced);
            Assert.Equal("Path bytes replacement", replacedInfo.Metadata.Title);
            Assert.Equal("Path bytes author", replacedInfo.Metadata.Author);
            Assert.Null(replacedInfo.Metadata.Subject);
            Assert.Null(replacedInfo.Metadata.Keywords);
            Assert.Equal(2, replacedInfo.PageCount);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void MetadataPathInputs_WriteToOutputStreamsForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-metadata-path-stream-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildTwoPagePdf());

            byte[] prefix = Encoding.ASCII.GetBytes("prefix");
            using var updateOutput = new MemoryStream();
            updateOutput.Write(prefix, 0, prefix.Length);

            PdfMetadataEditor.UpdateMetadata(
                inputPath,
                updateOutput,
                title: "Path stream title",
                keywords: "path,stream");

            byte[] updated = updateOutput.ToArray().Skip(prefix.Length).ToArray();
            Assert.Equal(prefix, updateOutput.ToArray().Take(prefix.Length).ToArray());
            PdfDocumentInfo updatedInfo = PdfInspector.Inspect(updated);
            Assert.Equal("Path stream title", updatedInfo.Metadata.Title);
            Assert.Equal("Original author", updatedInfo.Metadata.Author);
            Assert.Equal("Original subject", updatedInfo.Metadata.Subject);
            Assert.Equal("path,stream", updatedInfo.Metadata.Keywords);
            Assert.Equal(2, updatedInfo.PageCount);

            using var replaceOutput = new MemoryStream();
            replaceOutput.Write(prefix, 0, prefix.Length);

            PdfMetadataEditor.ReplaceMetadata(inputPath, replaceOutput, new PdfMetadata {
                Title = "Path stream replacement",
                Author = "Path stream author"
            });

            byte[] replaced = replaceOutput.ToArray().Skip(prefix.Length).ToArray();
            Assert.Equal(prefix, replaceOutput.ToArray().Take(prefix.Length).ToArray());
            PdfDocumentInfo replacedInfo = PdfInspector.Inspect(replaced);
            Assert.Equal("Path stream replacement", replacedInfo.Metadata.Title);
            Assert.Equal("Path stream author", replacedInfo.Metadata.Author);
            Assert.Null(replacedInfo.Metadata.Subject);
            Assert.Null(replacedInfo.Metadata.Keywords);
            Assert.Equal(2, replacedInfo.PageCount);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void MetadataEditor_RejectsNullInputs() {
        byte[] source = BuildTwoPagePdf();

        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.UpdateMetadata((byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.UpdateMetadata((Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.UpdateMetadata(new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.UpdateMetadata(source, null!, title: "x"));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.UpdateMetadata(source, new ReadOnlyStream(), title: "x"));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.UpdateMetadata(new MemoryStream(source), null!, title: "x"));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.UpdateMetadata(new MemoryStream(source), new ReadOnlyStream(), title: "x"));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.UpdateMetadata(null!, "out.pdf", title: "x"));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.UpdateMetadata("input.pdf", (string)null!, title: "x"));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.UpdateMetadata(" ", "out.pdf", title: "x"));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.UpdateMetadata("missing.pdf", " ", title: "x"));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.UpdateMetadata("input.pdf", (Stream)null!, title: "x"));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.UpdateMetadata("missing.pdf", new ReadOnlyStream(), title: "x"));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.UpdateMetadata(" ", new MemoryStream(), title: "x"));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.UpdateMetadataToBytes(null!, title: "x"));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.UpdateMetadataToBytes(" ", title: "x"));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadata(source, null!));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadata((byte[])null!, new PdfMetadata()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadata((Stream)null!, new PdfMetadata()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadata(new MemoryStream(source), null!));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.ReplaceMetadata(new WriteOnlyStream(), new PdfMetadata()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadata(source, null!, new PdfMetadata()));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.ReplaceMetadata(source, new ReadOnlyStream(), new PdfMetadata()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadata(new MemoryStream(source), null!, new PdfMetadata()));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.ReplaceMetadata(new MemoryStream(source), new ReadOnlyStream(), new PdfMetadata()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadata(null!, "out.pdf", new PdfMetadata()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadata("input.pdf", (string)null!, new PdfMetadata()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadata("input.pdf", "out.pdf", null!));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.ReplaceMetadata(" ", "out.pdf", new PdfMetadata()));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.ReplaceMetadata("missing.pdf", " ", new PdfMetadata()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadata("input.pdf", (Stream)null!, new PdfMetadata()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadata("input.pdf", new MemoryStream(), null!));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.ReplaceMetadata("missing.pdf", new ReadOnlyStream(), new PdfMetadata()));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.ReplaceMetadata(" ", new MemoryStream(), new PdfMetadata()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadataToBytes(null!, new PdfMetadata()));
        Assert.Throws<ArgumentNullException>(() => PdfMetadataEditor.ReplaceMetadataToBytes("input.pdf", null!));
        Assert.Throws<ArgumentException>(() => PdfMetadataEditor.ReplaceMetadataToBytes(" ", new PdfMetadata()));
    }

    [Fact]
    public void MetadataEditorPathOutputs_RejectDirectoryTargets() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-metadata-output-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputDirectory = Path.Combine(directory, "existing-output");

        try {
            Directory.CreateDirectory(outputDirectory);
            File.WriteAllBytes(inputPath, BuildTwoPagePdf());

            var ex = Assert.Throws<ArgumentException>(() => PdfMetadataEditor.UpdateMetadata(inputPath, outputDirectory, title: "Directory target"));
            Assert.Equal("outputPath", ex.ParamName);
            Assert.Contains("Output path refers to a directory; a file path is required.", ex.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static byte[] BuildTwoPagePdf() {
        var doc = PdfDocument.Create()
            .Meta(
                title: "Original title",
                author: "Original author",
                subject: "Original subject",
                keywords: "original,metadata");

        doc.Compose(compose => {
            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("First page"))));
            });

            compose.Page(page => {
                page.Size(new PageSize(300, 500));
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Second page"))));
            });
        });

        return doc.ToBytes();
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

    private sealed class WriteOnlyStream : MemoryStream {
        public override bool CanRead => false;
    }

    private sealed class ReadOnlyStream : MemoryStream {
        public override bool CanWrite => false;
    }
}

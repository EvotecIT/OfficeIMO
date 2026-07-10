using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Json;
using OfficeIMO.Word;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderDetectionTests {
    [Fact]
    public void DocumentReader_Detect_PrefersStrongPdfContentAndRestoresStreamPosition() {
        byte[] bytes = Encoding.ASCII.GetBytes("prefix%PDF-1.7\nbody");
        using var stream = new MemoryStream(bytes, writable: false);
        stream.Position = 6;

        ReaderDetectionResult detection = DocumentReader.Detect(
            stream,
            "report.txt",
            new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent });

        Assert.Equal(6, stream.Position);
        Assert.Equal(ReaderInputKind.Text, detection.ExtensionKind);
        Assert.Equal(ReaderInputKind.Pdf, detection.ContentKind);
        Assert.Equal(ReaderInputKind.Pdf, detection.Kind);
        Assert.Equal(ReaderDetectionConfidence.High, detection.ContentConfidence);
        Assert.True(detection.IsMismatch);
        Assert.Contains("signature:pdf", detection.Evidence);
        Assert.Equal("application/pdf", detection.MediaType);
    }

    [Fact]
    public void DocumentReader_Detect_ContentWhenUnknownDoesNotProbeKnownExtension() {
        byte[] bytes = Encoding.ASCII.GetBytes("%PDF-1.7\nbody");

        ReaderDetectionResult detection = DocumentReader.Detect(
            bytes,
            "report.txt",
            new ReaderDetectionOptions { Mode = ReaderDetectionMode.ContentWhenUnknown });

        Assert.Equal(ReaderInputKind.Text, detection.Kind);
        Assert.False(detection.ContentInspected);
        Assert.Equal(0, detection.InspectedBytes);
    }

    [Fact]
    public async Task OfficeDocumentReader_DetectAsync_UsesAsyncReadsForNonSeekableStream() {
        using var stream = new AsyncOnlyNonSeekableStream(Encoding.ASCII.GetBytes("%PDF-1.7\nbody"));

        ReaderDetectionResult detection = await OfficeDocumentReader.Default.DetectAsync(stream, "upload.blob");

        Assert.Equal(ReaderInputKind.Pdf, detection.Kind);
        Assert.Equal(ReaderDetectionConfidence.High, detection.Confidence);
        Assert.True(detection.ContentInspected);
        Assert.Contains("signature:pdf", detection.Evidence);
    }

    [Fact]
    public void OfficeDocumentReader_Detect_DoesNotOverrideHtmlExtensionWithAmbiguousFragment() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHtmlHandler()
            .Build();
        byte[] bytes = Encoding.UTF8.GetBytes("<div>HTML fragment</div>");

        ReaderDetectionResult detection = reader.Detect(bytes, "fragment.html");

        Assert.Equal(ReaderInputKind.Html, detection.ExtensionKind);
        Assert.Equal(ReaderInputKind.Xml, detection.ContentKind);
        Assert.Equal(ReaderDetectionConfidence.Low, detection.ContentConfidence);
        Assert.Equal(ReaderInputKind.Html, detection.Kind);
        Assert.False(detection.IsMismatch);
    }

    [Fact]
    public void DocumentReader_Detect_IdentifiesWordContainerWithoutWordExtension() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".blob");
        using var package = new MemoryStream();
        using (WordDocument document = WordDocument.Create(package)) {
            document.AddParagraph("Container detection");
            document.Save();
        }
        File.WriteAllBytes(path, package.ToArray());

        try {
            ReaderDetectionResult detection = DocumentReader.Detect(path);

            Assert.Equal(ReaderInputKind.Unknown, detection.ExtensionKind);
            Assert.Equal(ReaderInputKind.Word, detection.ContentKind);
            Assert.Equal(ReaderInputKind.Word, detection.Kind);
            Assert.Equal(ReaderDetectionConfidence.High, detection.Confidence);
            Assert.True(detection.ContainerInspected);
            Assert.Contains("container:word/document.xml", detection.Evidence);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task DocumentReader_DetectAsync_IdentifiesWordContainerWithoutWordExtension() {
        using var package = new MemoryStream();
        using (WordDocument document = WordDocument.Create(package)) {
            document.AddParagraph("Async container detection");
            document.Save();
        }

        ReaderDetectionResult detection = await DocumentReader.DetectAsync(package.ToArray(), "document.blob");

        Assert.Equal(ReaderInputKind.Word, detection.Kind);
        Assert.Equal(ReaderDetectionConfidence.High, detection.Confidence);
        Assert.True(detection.ContainerInspected);
        Assert.Contains("container:word/document.xml", detection.Evidence);
    }

    [Fact]
    public void OfficeDocumentReader_RoutesUnknownExtensionToUniqueContentHandler() {
        const string handlerId = "officeimo.tests.detection.pdf";
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".blob");
        File.WriteAllText(path, "%PDF-1.7\nbody");

        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddHandler(new ReaderHandlerRegistration {
                    Id = handlerId,
                    Kind = ReaderInputKind.Pdf,
                    Extensions = new[] { ".pdf" },
                    ReadDocumentPath = (sourcePath, options, cancellationToken) => new OfficeDocumentReadResult {
                        Kind = ReaderInputKind.Pdf,
                        Source = new OfficeDocumentSource { Path = sourcePath },
                        CapabilitiesUsed = new[] { handlerId },
                        Chunks = new[] {
                            new ReaderChunk {
                                Id = "detected-pdf:0001",
                                Kind = ReaderInputKind.Pdf,
                                Text = "content-routed-pdf"
                            }
                        }
                    }
                }, replaceExisting: true)
                .Build();

            OfficeDocumentReadResult result = reader.ReadDocument(path);

            Assert.Equal("content-routed-pdf", Assert.Single(result.Chunks).Text);
            Assert.Contains(handlerId, result.CapabilitiesUsed);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "input-kind-detected");
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsStructuredMismatchDiagnostic() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, "# Detected Markdown\n\nBody");

        try {
            OfficeDocumentReadResult result = DocumentReader.ReadDocument(
                path,
                new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });

            Assert.Equal(ReaderInputKind.Markdown, result.Kind);
            OfficeDocumentDiagnostic diagnostic = Assert.Single(
                result.Diagnostics,
                item => item.Code == "input-kind-mismatch");
            Assert.Equal(OfficeDocumentDiagnosticCategory.Detection, diagnostic.Category);
            Assert.Equal(OfficeDocumentDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.True(diagnostic.IsRecoverable);
            Assert.Equal("Text", diagnostic.Attributes["extensionKind"]);
            Assert.Equal("Markdown", diagnostic.Attributes["contentKind"]);

            using JsonDocument json = JsonDocument.Parse(result.ToJson());
            JsonElement jsonDiagnostic = json.RootElement.GetProperty("diagnostics")[0];
            Assert.Equal("Detection", jsonDiagnostic.GetProperty("category").GetString());
            Assert.Equal("officeimo.reader.detection", jsonDiagnostic.GetProperty("source").GetString());
            Assert.True(jsonDiagnostic.GetProperty("isRecoverable").GetBoolean());
            Assert.Equal("Markdown", jsonDiagnostic.GetProperty("attributes").GetProperty("contentKind").GetString());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsDetectedKindDiagnosticForUnknownExtension() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".blob");
        File.WriteAllText(path, "# Detected Markdown\n\nBody");

        try {
            OfficeDocumentReadResult result = DocumentReader.ReadDocument(path);

            Assert.Equal(ReaderInputKind.Markdown, result.Kind);
            OfficeDocumentDiagnostic diagnostic = Assert.Single(
                result.Diagnostics,
                item => item.Code == "input-kind-detected");
            Assert.Equal(OfficeDocumentDiagnosticSeverity.Information, diagnostic.Severity);
            Assert.Equal(OfficeDocumentDiagnosticCategory.Detection, diagnostic.Category);
            Assert.Equal("Markdown", diagnostic.Attributes["contentKind"]);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task OfficeDocumentReader_ReadDocumentAsync_RoutesDetectedKindToAsyncOnlyHandler() {
        const string handlerId = "officeimo.tests.detection.async-pdf";
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".blob");
        File.WriteAllText(path, "%PDF-1.7\nbody");

        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddHandler(new ReaderHandlerRegistration {
                    Id = handlerId,
                    Kind = ReaderInputKind.Pdf,
                    Extensions = new[] { ".pdf" },
                    ReadDocumentPathAsync = async (sourcePath, options, cancellationToken) => {
                        await Task.Yield();
                        return new OfficeDocumentReadResult {
                            Kind = ReaderInputKind.Pdf,
                            Source = new OfficeDocumentSource { Path = sourcePath },
                            CapabilitiesUsed = new[] { handlerId }
                        };
                    }
                }, replaceExisting: true)
                .Build();

            OfficeDocumentReadResult result = await reader.ReadDocumentAsync(path);

            Assert.Contains(handlerId, result.CapabilitiesUsed);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "input-kind-detected");
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OfficeDocumentReader_MapsAdapterWarningToStableParsingDiagnostic() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddJsonHandler()
            .Build();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("{ invalid json"));

        OfficeDocumentReadResult result = reader.ReadDocument(stream, "broken.json");

        OfficeDocumentDiagnostic diagnostic = Assert.Single(
            result.Diagnostics,
            item => item.Code == "parse-failed");
        Assert.Equal(OfficeDocumentDiagnosticCategory.Parsing, diagnostic.Category);
        Assert.Equal("officeimo.reader", diagnostic.Source);
        Assert.True(diagnostic.IsRecoverable);
    }

    private sealed class AsyncOnlyNonSeekableStream : Stream {
        private readonly byte[] _bytes;
        private int _position;

        public AsyncOnlyNonSeekableStream(byte[] bytes) {
            _bytes = bytes;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => _position;
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) {
            throw new InvalidOperationException("Synchronous reads are not allowed.");
        }

        public override Task<int> ReadAsync(
            byte[] buffer,
            int offset,
            int count,
            CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            int available = Math.Min(count, _bytes.Length - _position);
            if (available > 0) {
                Array.Copy(_bytes, _position, buffer, offset, available);
                _position += available;
            }
            return Task.FromResult(available);
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}

using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Json;
using OfficeIMO.Word;
using System.IO.Compression;
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

        ReaderDetectionResult detection = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Detect(
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

        ReaderDetectionResult detection = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Detect(
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

        ReaderDetectionResult detection = await OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectAsync(stream, "upload.blob");

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
            ReaderDetectionResult detection = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Detect(path);

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

        ReaderDetectionResult detection = await OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectAsync(package.ToArray(), "document.blob");

        Assert.Equal(ReaderInputKind.Word, detection.Kind);
        Assert.Equal(ReaderDetectionConfidence.High, detection.Confidence);
        Assert.True(detection.ContainerInspected);
        Assert.Contains("container:word/document.xml", detection.Evidence);
    }

    [Theory]
    [InlineData("application/vnd.oasis.opendocument.text")]
    [InlineData("application/vnd.oasis.opendocument.spreadsheet")]
    [InlineData("application/vnd.oasis.opendocument.presentation")]
    public void DocumentReader_Detect_IdentifiesOpenDocumentMimeTypes(string mediaType) {
        byte[] package = CreateStoredZip(
            ("mimetype", mediaType),
            ("content.xml", "<document />"));
        Assert.Equal(0, BitConverter.ToUInt16(package, 8));

        ReaderDetectionResult detection = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Detect(package, "document.blob");

        Assert.Equal(ReaderInputKind.OpenDocument, detection.Kind);
        Assert.Equal(mediaType, detection.MediaType);
        Assert.Contains("container:opendocument-mimetype", detection.Evidence);
    }

    [Fact]
    public async Task DocumentReader_DetectAsync_IdentifiesOpenDocumentMimeType() {
        const string mediaType = "application/vnd.oasis.opendocument.text";
        byte[] package = CreateStoredZip(
            ("mimetype", mediaType),
            ("content.xml", "<document />"));
        Assert.Equal(0, BitConverter.ToUInt16(package, 8));

        ReaderDetectionResult detection = await OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectAsync(package, "document.blob");

        Assert.Equal(ReaderInputKind.OpenDocument, detection.Kind);
        Assert.Equal(mediaType, detection.MediaType);
        Assert.Contains("container:opendocument-mimetype", detection.Evidence);
    }

    [Fact]
    public void DocumentReader_Detect_InspectsEntriesAfterDataDescriptors() {
        byte[] package = CreateStreamingZip(
            ("first.bin", "descriptor payload"),
            ("word/document.xml", "<document />"));
        Assert.NotEqual(0, BitConverter.ToUInt16(package, 6) & 0x0008);

        ReaderDetectionResult detection = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Detect(package, "document.blob");

        Assert.Equal(ReaderInputKind.Word, detection.Kind);
        Assert.Equal(ReaderDetectionConfidence.High, detection.Confidence);
        Assert.Contains("container:word/document.xml", detection.Evidence);
    }

    [Fact]
    public async Task DocumentReader_DetectAsync_MatchesContainerMarkersOnlyFromEntryNames() {
        byte[] package = CreateStreamingZip(
            ("notes.txt", "payload mentions word/document.xml but is not an Office package"),
            ("prefixword/document.xmlsuffix", "partial entry-name match"));

        ReaderDetectionResult detection = await OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectAsync(package, "document.blob");

        Assert.Equal(ReaderInputKind.Zip, detection.Kind);
        Assert.Equal(ReaderInputKind.Zip, detection.ContentKind);
        Assert.DoesNotContain("container:word/document.xml", detection.Evidence);
    }

    [Fact]
    public void DocumentReader_Detect_RecognizesUtf16JsonBeforeBinaryHeuristics() {
        byte[] bytes = Encoding.Unicode.GetPreamble()
            .Concat(Encoding.Unicode.GetBytes("{\"name\":\"OfficeIMO\"}"))
            .ToArray();

        ReaderDetectionResult detection = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Detect(bytes, "document.blob");

        Assert.Equal(ReaderInputKind.Json, detection.Kind);
        Assert.Equal("application/json", detection.MediaType);
    }

    [Theory]
    [InlineData("records.csv", ReaderInputKind.Csv, "text/csv")]
    [InlineData("records.tsv", ReaderInputKind.Csv, "text/tab-separated-values")]
    [InlineData("document.json", ReaderInputKind.Json, "application/json")]
    [InlineData("document.xml", ReaderInputKind.Xml, "application/xml")]
    [InlineData("document.yaml", ReaderInputKind.Yaml, "application/yaml")]
    public void DocumentReader_Detect_PreservesExtensionSpecificMediaTypes(string sourceName, ReaderInputKind expectedKind, string expectedMediaType) {
        ReaderDetectionResult detection = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Detect(
            Array.Empty<byte>(),
            sourceName,
            new ReaderDetectionOptions { Mode = ReaderDetectionMode.ExtensionOnly });

        Assert.Equal(expectedKind, detection.Kind);
        Assert.Equal(expectedMediaType, detection.MediaType);
    }

    [Fact]
    public void DocumentReader_ReadDocument_EnforcesPathLimitBeforeContentDetection() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".blob");
        File.WriteAllBytes(path, new byte[257]);

        try {
            IOException exception = Assert.Throws<IOException>(() => OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(
                path,
                new ReaderOptions {
                    DetectionMode = ReaderDetectionMode.PreferContent,
                    MaxInputBytes = 256
                }));

            Assert.Contains("MaxInputBytes", exception.Message, StringComparison.Ordinal);
        } finally {
            File.Delete(path);
        }
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
            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(
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
            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(path);

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

    private static byte[] CreateStreamingZip(params (string Name, string Content)[] entries) {
        using var buffer = new MemoryStream();
        using (var output = new NonSeekableWriteStream(buffer))
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach ((string name, string content) in entries) {
                ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
                using Stream entryStream = entry.Open();
                byte[] bytes = Encoding.UTF8.GetBytes(content);
                entryStream.Write(bytes, 0, bytes.Length);
            }
        }
        return buffer.ToArray();
    }

    private static byte[] CreateStoredZip(params (string Name, string Content)[] entries) {
        using var buffer = new MemoryStream();
        using var writer = new BinaryWriter(buffer, Encoding.UTF8, leaveOpen: true);
        var records = new List<StoredZipEntry>(entries.Length);

        foreach ((string name, string content) in entries) {
            byte[] nameBytes = Encoding.UTF8.GetBytes(name);
            byte[] contentBytes = Encoding.UTF8.GetBytes(content);
            uint crc32 = CalculateCrc32(contentBytes);
            uint localHeaderOffset = checked((uint)buffer.Position);

            writer.Write(0x04034B50U);
            writer.Write((ushort)20);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write(crc32);
            writer.Write((uint)contentBytes.Length);
            writer.Write((uint)contentBytes.Length);
            writer.Write((ushort)nameBytes.Length);
            writer.Write((ushort)0);
            writer.Write(nameBytes);
            writer.Write(contentBytes);
            records.Add(new StoredZipEntry(nameBytes, contentBytes.Length, crc32, localHeaderOffset));
        }

        uint centralDirectoryOffset = checked((uint)buffer.Position);
        foreach (StoredZipEntry record in records) {
            writer.Write(0x02014B50U);
            writer.Write((ushort)20);
            writer.Write((ushort)20);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write(record.Crc32);
            writer.Write((uint)record.ContentLength);
            writer.Write((uint)record.ContentLength);
            writer.Write((ushort)record.NameBytes.Length);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write((ushort)0);
            writer.Write(0U);
            writer.Write(record.LocalHeaderOffset);
            writer.Write(record.NameBytes);
        }

        uint centralDirectorySize = checked((uint)buffer.Position - centralDirectoryOffset);
        writer.Write(0x06054B50U);
        writer.Write((ushort)0);
        writer.Write((ushort)0);
        writer.Write((ushort)records.Count);
        writer.Write((ushort)records.Count);
        writer.Write(centralDirectorySize);
        writer.Write(centralDirectoryOffset);
        writer.Write((ushort)0);
        writer.Flush();
        return buffer.ToArray();
    }

    private static uint CalculateCrc32(byte[] bytes) {
        uint crc = uint.MaxValue;
        foreach (byte value in bytes) {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320U : crc >> 1;
            }
        }
        return ~crc;
    }

    private sealed class StoredZipEntry {
        public StoredZipEntry(byte[] nameBytes, int contentLength, uint crc32, uint localHeaderOffset) {
            NameBytes = nameBytes;
            ContentLength = contentLength;
            Crc32 = crc32;
            LocalHeaderOffset = localHeaderOffset;
        }

        public byte[] NameBytes { get; }
        public int ContentLength { get; }
        public uint Crc32 { get; }
        public uint LocalHeaderOffset { get; }
    }

    private sealed class NonSeekableWriteStream : Stream {
        private readonly Stream _inner;

        public NonSeekableWriteStream(Stream inner) {
            _inner = inner;
        }

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            // .NET Framework's streaming ZipArchive reads the forward write position even when CanSeek is false.
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Flush();
            base.Dispose(disposing);
        }
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

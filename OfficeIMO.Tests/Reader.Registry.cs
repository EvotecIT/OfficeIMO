using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Zip;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

[CollectionDefinition("ReaderRegistryNonParallel", DisableParallelization = true)]
public sealed class ReaderRegistryNonParallelCollection;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderRegistryTests {
    [Fact]
    public void DocumentReader_GetCapabilities_IncludesBuiltInHandlers() {
        var capabilities = DocumentReader.GetCapabilities();

        Assert.NotEmpty(capabilities);
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.word");
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.excel");
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.powerpoint");
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.markdown");
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.pdf");
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.text");
    }

    [Fact]
    public void DocumentReader_RegisterHandler_UsesCustomPathReader() {
        const string handlerId = "officeimo.tests.custom.demo";
        const string extension = ".demoix";

        var file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
        try {
            DocumentReader.UnregisterHandler(handlerId);

            DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
                Id = handlerId,
                DisplayName = "Test custom handler",
                Kind = ReaderInputKind.Text,
                Extensions = new[] { extension },
                ReadPath = (path, options, ct) => new[] {
                    new ReaderChunk {
                        Id = "custom-0001",
                        Kind = ReaderInputKind.Text,
                        Location = new ReaderLocation {
                            Path = path,
                            BlockIndex = 0
                        },
                        Text = "custom-handler-output"
                    }
                }
            });

            File.WriteAllText(file, "input");

            var kind = DocumentReader.DetectKind(file);
            Assert.Equal(ReaderInputKind.Text, kind);

            var chunks = DocumentReader.Read(file).ToList();
            Assert.Single(chunks);
            Assert.Equal("custom-handler-output", chunks[0].Text);

            var customCapabilities = DocumentReader.GetCapabilities(includeBuiltIn: false, includeCustom: true);
            Assert.Contains(customCapabilities, c => c.Id == handlerId && c.Extensions.Contains(extension, StringComparer.OrdinalIgnoreCase));
        } finally {
            DocumentReader.UnregisterHandler(handlerId);
            if (File.Exists(file)) File.Delete(file);
        }
    }

    [Fact]
    public void DocumentReader_RegisterHandler_WithoutReplaceExisting_RejectsBuiltInCollision() {
        const string handlerId = "officeimo.tests.custom.markdown";

        DocumentReader.UnregisterHandler(handlerId);
        try {
            Assert.Throws<InvalidOperationException>(() => DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
                Id = handlerId,
                Extensions = new[] { ".md" },
                Kind = ReaderInputKind.Markdown,
                ReadPath = (path, options, ct) => Array.Empty<ReaderChunk>()
            }));
        } finally {
            DocumentReader.UnregisterHandler(handlerId);
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_RegisterAndUnregister() {
        try {
            DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler(replaceExisting: true);
            DocumentReaderZipRegistrationExtensions.RegisterZipHandler(replaceExisting: true);
            DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(replaceExisting: true);

            var capabilities = DocumentReader.GetCapabilities();
            var epubCapability = Assert.Single(capabilities, c => c.Id == DocumentReaderEpubRegistrationExtensions.HandlerId);
            var zipCapability = Assert.Single(capabilities, c => c.Id == DocumentReaderZipRegistrationExtensions.HandlerId);
            var htmlCapability = Assert.Single(capabilities, c => c.Id == DocumentReaderHtmlRegistrationExtensions.HandlerId);

            Assert.True(epubCapability.SupportsPath);
            Assert.True(epubCapability.SupportsStream);
            Assert.True(zipCapability.SupportsPath);
            Assert.True(zipCapability.SupportsStream);
            Assert.True(htmlCapability.SupportsPath);
            Assert.True(htmlCapability.SupportsStream);

            Assert.Equal(ReaderInputKind.Unknown, DocumentReader.DetectKind("book.epub"));
            Assert.Equal(ReaderInputKind.Unknown, DocumentReader.DetectKind("archive.zip"));
            Assert.Equal(ReaderInputKind.Unknown, DocumentReader.DetectKind("index.html"));
        } finally {
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_DispatchesZipStream() {
        try {
            DocumentReaderZipRegistrationExtensions.RegisterZipHandler(replaceExisting: true);

            using var stream = new MemoryStream();
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteTextEntry(archive, "docs/readme.md", "# Stream ZIP" + Environment.NewLine + Environment.NewLine + "Body from zip stream.");
            }

            stream.Position = 0;
            var chunks = DocumentReader.Read(stream, "bundle.zip").ToList();

            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Markdown &&
                (c.Location.Path?.Contains("bundle.zip::docs/readme.md", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("Body from zip stream.", StringComparison.Ordinal) ?? false));
        } finally {
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_DispatchesEpubStream() {
        try {
            DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler(replaceExisting: true);

            var bytes = BuildSimpleEpubBytes();
            using var stream = new MemoryStream(bytes, writable: false);
            var chunks = DocumentReader.Read(stream, "book.epub").ToList();

            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Unknown &&
                (c.Location.Path?.Contains("book.epub::OEBPS/chapter.xhtml", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("EPUB stream body text.", StringComparison.Ordinal) ?? false));
        } finally {
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
        }
    }

    private static byte[] BuildSimpleEpubBytes() {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            WriteTextEntry(archive, "META-INF/container.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
                "<rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles>" +
                "</container>");

            WriteTextEntry(archive, "OEBPS/content.opf",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\">" +
                "<manifest><item id=\"ch1\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/></manifest>" +
                "<spine><itemref idref=\"ch1\"/></spine>" +
                "</package>");

            WriteTextEntry(archive, "OEBPS/chapter.xhtml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><p>EPUB stream body text.</p></body></html>");
        }

        return ms.ToArray();
    }

    private static void WriteTextEntry(ZipArchive archive, string path, string content, CompressionLevel compressionLevel = CompressionLevel.Optimal) {
        var entry = archive.CreateEntry(path, compressionLevel);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), 4096, leaveOpen: false);
        writer.Write(content);
    }
}

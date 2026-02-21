using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Zip;
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
            Assert.Contains(capabilities, c => c.Id == DocumentReaderEpubRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderZipRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderHtmlRegistrationExtensions.HandlerId);

            Assert.Equal(ReaderInputKind.Unknown, DocumentReader.DetectKind("book.epub"));
            Assert.Equal(ReaderInputKind.Unknown, DocumentReader.DetectKind("archive.zip"));
            Assert.Equal(ReaderInputKind.Unknown, DocumentReader.DetectKind("index.html"));
        } finally {
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
        }
    }
}

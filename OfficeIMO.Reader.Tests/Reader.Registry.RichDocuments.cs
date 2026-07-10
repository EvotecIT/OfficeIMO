using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderRegistryTests {
    [Fact]
    public void DocumentReader_RegisterHandler_RichDocumentOnly_ProjectsChunksAndPreservesEnvelope() {
        const string handlerId = "officeimo.tests.custom.rich";
        const string extension = ".richix";
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

        try {
            DocumentReader.UnregisterHandler(handlerId);
            DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
                Id = handlerId,
                DisplayName = "Rich document test handler",
                Kind = ReaderInputKind.Text,
                Extensions = new[] { extension },
                ReadDocumentPath = (path, options, ct) => new OfficeDocumentReadResult {
                    Kind = ReaderInputKind.Text,
                    Source = new OfficeDocumentSource { Path = path },
                    CapabilitiesUsed = new[] { handlerId, "officeimo.tests.rich-envelope" },
                    Chunks = new[] {
                        new ReaderChunk {
                            Id = "rich-0001",
                            Kind = ReaderInputKind.Text,
                            Location = new ReaderLocation { Path = path, BlockIndex = 0 },
                            Text = "rich-document-output"
                        }
                    },
                    Links = new[] {
                        new OfficeDocumentLink {
                            Id = "rich-link-0001",
                            Kind = "external",
                            Uri = "https://example.test/rich",
                            Location = new ReaderLocation { Path = path }
                        }
                    }
                }
            });
            File.WriteAllText(file, "input");

            ReaderChunk chunk = Assert.Single(DocumentReader.Read(file));
            Assert.Equal("rich-document-output", chunk.Text);

            OfficeDocumentReadResult result = DocumentReader.ReadDocument(file);
            Assert.Contains("officeimo.tests.rich-envelope", result.CapabilitiesUsed);
            Assert.Equal("https://example.test/rich", Assert.Single(result.Links).Uri);

            ReaderHandlerCapability capability = Assert.Single(
                DocumentReader.GetCapabilities(includeBuiltIn: false, includeCustom: true),
                item => item.Id == handlerId);
            Assert.True(capability.SupportsPath);
            Assert.False(capability.SupportsStream);
            Assert.True(capability.SupportsDocumentPath);
            Assert.False(capability.SupportsDocumentStream);
        } finally {
            DocumentReader.UnregisterHandler(handlerId);
            if (File.Exists(file)) File.Delete(file);
        }
    }
}

using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderRegistryTests {
    [Fact]
    public void OfficeDocumentReader_HandlerDefaultLimit_ReportsTheRegisteredPathOnlyHandler() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "officeimo.tests.path-only-email",
                Kind = ReaderInputKind.Email,
                Extensions = new[] { ".eml" },
                DefaultMaxInputBytes = 16,
                ReadPath = (path, options, cancellationToken) => Array.Empty<ReaderChunk>()
            }, replaceExisting: true)
            .Build();

        Assert.Equal(16, reader.GetHandlerDefaultMaxInputBytes("message.eml"));
        ReaderHandlerCapability capability = Assert.Single(reader.GetCapabilities());
        Assert.True(capability.SupportsPath);
        Assert.False(capability.SupportsStream);
    }

    [Fact]
    public void OfficeDocumentReader_HandlerDefaultLimit_HonorsNullableStreamOverride() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "officeimo.tests.unbounded-email-stream",
                Kind = ReaderInputKind.Email,
                Extensions = new[] { ".eml" },
                ReadStream = (stream, sourceName, options, cancellationToken) => Array.Empty<ReaderChunk>()
            }, replaceExisting: true)
            .Build();

        Assert.Null(reader.GetHandlerDefaultMaxInputBytes("message.eml"));
    }

    [Fact]
    public void OfficeDocumentReader_RichDocumentHandler_ProjectsChunksAndPreservesEnvelope() {
        const string handlerId = "officeimo.tests.custom.rich";
        const string extension = ".richix";
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddHandler(new ReaderHandlerRegistration {
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
                })
                .Build();
            File.WriteAllText(file, "input");

            ReaderChunk chunk = Assert.Single(reader.Read(file));
            Assert.Equal("rich-document-output", chunk.Text);

            OfficeDocumentReadResult result = reader.ReadDocument(file);
            Assert.Contains("officeimo.tests.rich-envelope", result.CapabilitiesUsed);
            Assert.Equal("https://example.test/rich", Assert.Single(result.Links).Uri);

            ReaderHandlerCapability capability = Assert.Single(
                reader.GetCapabilities(),
                item => item.Id == handlerId);
            Assert.True(capability.SupportsPath);
            Assert.False(capability.SupportsStream);
            Assert.True(capability.SupportsDocumentPath);
            Assert.False(capability.SupportsDocumentStream);
        } finally {
            if (File.Exists(file)) File.Delete(file);
        }
    }

    [Fact]
    public void OfficeDocumentReader_RichStreamHandler_EnforcesNonSeekableInputLimitBeforeDispatch() {
        const string handlerId = "officeimo.tests.custom.rich.limit";
        const string extension = ".richlimitix";
        bool handlerInvoked = false;

        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = handlerId,
                Kind = ReaderInputKind.Text,
                Extensions = new[] { extension },
                ReadDocumentStream = (stream, sourceName, options, cancellationToken) => {
                    handlerInvoked = true;
                    return new OfficeDocumentReadResult { Kind = ReaderInputKind.Text };
                }
            })
            .Build();
        using var stream = new NonSeekableReadStream(new byte[128]);

        IOException exception = Assert.Throws<IOException>(() => reader.ReadDocument(
            stream,
            "sample" + extension,
            new ReaderOptions { MaxInputBytes = 16 }));

        Assert.Contains("Input exceeds MaxInputBytes", exception.Message, StringComparison.Ordinal);
        Assert.False(handlerInvoked);
    }
}

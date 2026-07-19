using OfficeIMO.Reader;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderMarkdownConversionTests {
    [Fact]
    public void ConvertToMarkdown_PathProjectsRichResultMarkdown() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path, "# Policy\n\nPortable content.\n");

            OfficeDocumentReadResult document = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(path);
            string markdown = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ConvertToMarkdown(path);

            Assert.Equal(document.Markdown ?? string.Empty, markdown);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ConvertToMarkdown_StreamAndBytesProjectRichResultMarkdown() {
        byte[] bytes = Encoding.UTF8.GetBytes("# Stream\n\nContent.\n");
        using var expectedStream = new MemoryStream(bytes, writable: false);
        using var conversionStream = new MemoryStream(bytes, writable: false);
        OfficeDocumentReadResult document = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(expectedStream, "sample.md");

        string streamMarkdown = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ConvertToMarkdown(conversionStream, "sample.md");
        string bytesMarkdown = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ConvertToMarkdown(bytes, "sample.md");

        Assert.Equal(document.Markdown ?? string.Empty, streamMarkdown);
        Assert.Equal(document.Markdown ?? string.Empty, bytesMarkdown);
        Assert.True(conversionStream.CanRead);
    }

    [Fact]
    public async Task ConvertToMarkdownAsync_PathStreamAndBytesProjectRichResultMarkdown() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        byte[] bytes = Encoding.UTF8.GetBytes("# Async\n\nContent.\n");
        try {
            File.WriteAllBytes(path, bytes);
            OfficeDocumentReadResult document = await OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocumentAsync(path);
            using var stream = new MemoryStream(bytes, writable: false);

            string pathMarkdown = await OfficeIMO.Reader.Tests.ReaderTestReaders.All.ConvertToMarkdownAsync(path);
            string streamMarkdown = await OfficeIMO.Reader.Tests.ReaderTestReaders.All.ConvertToMarkdownAsync(stream, "sample.md");
            string bytesMarkdown = await OfficeIMO.Reader.Tests.ReaderTestReaders.All.ConvertToMarkdownAsync(bytes, "sample.md");

            string expected = document.Markdown ?? string.Empty;
            Assert.Equal(expected, pathMarkdown);
            Assert.Equal(expected, streamMarkdown);
            Assert.Equal(expected, bytesMarkdown);
            Assert.True(stream.CanRead);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ConvertToMarkdown_UsesConfiguredRichHandler() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "officeimo.tests.markdown-conversion",
                Kind = ReaderInputKind.Text,
                Extensions = new[] { ".convertix" },
                ReadDocumentStream = (stream, sourceName, options, cancellationToken) => new OfficeDocumentReadResult {
                    Kind = ReaderInputKind.Text,
                    Markdown = "# Configured handler"
                }
            })
            .Build();

        string markdown = reader.ConvertToMarkdown(new byte[] { 1 }, "sample.convertix");

        Assert.Equal("# Configured handler", markdown);
    }

    [Fact]
    public async Task ConvertToMarkdownAsync_UsesNativeAsyncRichHandler() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "officeimo.tests.markdown-conversion-async",
                Kind = ReaderInputKind.Text,
                Extensions = new[] { ".asyncconvertix" },
                ReadDocumentStreamAsync = (stream, sourceName, options, cancellationToken) => Task.FromResult(
                    new OfficeDocumentReadResult {
                        Kind = ReaderInputKind.Text,
                        Markdown = "# Native async handler"
                    })
            })
            .Build();

        string markdown = await reader.ConvertToMarkdownAsync(new byte[] { 1 }, "sample.asyncconvertix");

        Assert.Equal("# Native async handler", markdown);
    }

    [Fact]
    public void ConvertToMarkdown_ReturnsEmptyStringWhenRichResultHasNoMarkdown() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "officeimo.tests.markdown-conversion-empty",
                Kind = ReaderInputKind.Text,
                Extensions = new[] { ".emptyconvertix" },
                ReadDocumentStream = (stream, sourceName, options, cancellationToken) => new OfficeDocumentReadResult {
                    Kind = ReaderInputKind.Text
                }
            })
            .Build();

        string markdown = reader.ConvertToMarkdown(new byte[] { 1 }, "sample.emptyconvertix");

        Assert.Equal(string.Empty, markdown);
    }
}

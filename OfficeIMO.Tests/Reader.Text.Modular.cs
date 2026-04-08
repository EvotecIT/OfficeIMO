using OfficeIMO.Reader;
using OfficeIMO.Reader.Text;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderTextModularTests {
    [Fact]
    public void DocumentReaderText_ReadStructuredText_ParsesJsonIntoStructuredChunks() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-json-" + Guid.NewGuid().ToString("N") + ".json");
        try {
            File.WriteAllText(path,
                "{\n" +
                "  \"user\": {\n" +
                "    \"name\": \"Alice\",\n" +
                "    \"roles\": [\"admin\", \"ops\"],\n" +
                "    \"active\": true\n" +
                "  }\n" +
                "}");

            var chunks = DocumentReaderTextExtensions.ReadStructuredText(
                path,
                structuredOptions: new StructuredTextReadOptions {
                    JsonChunkRows = 2,
                    IncludeJsonMarkdown = true
                }).ToList();

            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Json, c.Kind));
            Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("$.user.name", StringComparison.Ordinal));
            Assert.Contains(chunks, c => c.Tables != null && c.Tables.Count > 0 && c.Tables[0].Columns.Contains("Path", StringComparer.Ordinal));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderText_ReadStructuredText_ParsesXmlIntoStructuredChunks() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path,
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<catalog><book id=\"b1\"><title>One</title></book><book id=\"b2\"><title>Two</title></book></catalog>");

            var chunks = DocumentReaderTextExtensions.ReadStructuredText(
                path,
                structuredOptions: new StructuredTextReadOptions {
                    XmlChunkRows = 2,
                    IncludeXmlMarkdown = true
                }).ToList();

            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Xml, c.Kind));
            Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("catalog[1]/book[1]/@id", StringComparison.Ordinal));
            Assert.Contains(chunks, c => c.Tables != null && c.Tables.Count > 0 && c.Tables[0].Columns.Contains("Type", StringComparer.Ordinal));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderText_ReadStructuredTextStream_ParsesJsonIntoStructuredChunks() {
        var json =
            "{\n" +
            "  \"agent\": {\n" +
            "    \"name\": \"OfficeIMO\",\n" +
            "    \"version\": 1\n" +
            "  }\n" +
            "}";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(json), writable: false);

        var chunks = DocumentReaderTextExtensions.ReadStructuredText(
            stream,
            sourceName: "agent.json",
            structuredOptions: new StructuredTextReadOptions {
                JsonChunkRows = 2,
                IncludeJsonMarkdown = true
            }).ToList();

        Assert.NotEmpty(chunks);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Json, c.Kind));
        Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("$.agent.name", StringComparison.Ordinal));
        Assert.Contains(chunks, c => c.Tables != null && c.Tables.Count > 0 && c.Tables[0].Columns.Contains("Path", StringComparer.Ordinal));
        Assert.All(chunks, c => {
            Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
            Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
            Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
            Assert.Equal(stream.Length, c.SourceLengthBytes);
            Assert.Null(c.SourceLastWriteUtc);
        });
    }

    [Fact]
    public void DocumentReaderText_ReadStructuredTextStream_ParsesCsvIntoStructuredChunks() {
        var csv =
            "Name,Role\n" +
            "Alice,Admin\n" +
            "Bob,Ops\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv), writable: false);

        var chunks = DocumentReaderTextExtensions.ReadStructuredText(
            stream,
            sourceName: "users.csv",
            structuredOptions: new StructuredTextReadOptions {
                CsvChunkRows = 1,
                IncludeCsvMarkdown = true
            }).ToList();

        Assert.NotEmpty(chunks);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Csv, c.Kind));
        Assert.Contains(chunks, c => (c.Location.Path ?? string.Empty).Contains("users.csv", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(chunks, c => c.Tables != null && c.Tables.Count > 0 && c.Tables[0].Columns.Contains("Name", StringComparer.Ordinal));
        Assert.All(chunks, c => {
            Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
            Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
            Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
            Assert.Equal(stream.Length, c.SourceLengthBytes);
            Assert.Null(c.SourceLastWriteUtc);
        });
    }

    [Fact]
    public void DocumentReaderText_ReadStructuredTextStream_NonSeekable_EnforcesMaxInputBytes() {
        var json =
            "{\n" +
            "  \"agent\": {\n" +
            "    \"name\": \"OfficeIMO\"\n" +
            "  }\n" +
            "}";
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(json));

        var ex = Assert.Throws<IOException>(() => DocumentReaderTextExtensions.ReadStructuredText(
            stream,
            sourceName: "agent.json",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 },
            structuredOptions: new StructuredTextReadOptions {
                JsonChunkRows = 2,
                IncludeJsonMarkdown = true
            }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderText_ReadStructuredText_EmitsWarningForMalformedXml() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-bad-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path, "<root><broken></root>");

            var warningChunk = Assert.Single(
                DocumentReaderTextExtensions.ReadStructuredText(path, readerOptions: new ReaderOptions { ComputeHashes = true }),
                c => c.Kind == ReaderInputKind.Xml &&
                     (c.Warnings?.Any(w => w.Contains("XML parse error", StringComparison.OrdinalIgnoreCase)) ?? false));

            Assert.False(string.IsNullOrWhiteSpace(warningChunk.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(warningChunk.SourceHash));
            Assert.False(string.IsNullOrWhiteSpace(warningChunk.ChunkHash));
            Assert.True(warningChunk.TokenEstimate.HasValue && warningChunk.TokenEstimate.Value >= 1);
            Assert.True(warningChunk.SourceLengthBytes.HasValue && warningChunk.SourceLengthBytes.Value > 0);
            Assert.True(warningChunk.SourceLastWriteUtc.HasValue);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderText_ReadStructuredText_PreservesQualifiedXmlNamesInPaths() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-structured-ns-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path,
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<root xmlns:a=\"urn:alpha\" xmlns:b=\"urn:beta\">" +
                "<a:item code=\"A1\">One</a:item>" +
                "<b:item b:code=\"B1\">Two</b:item>" +
                "</root>");

            var chunks = DocumentReaderTextExtensions.ReadStructuredText(
                path,
                structuredOptions: new StructuredTextReadOptions {
                    XmlChunkRows = 20,
                    IncludeXmlMarkdown = true
                }).ToList();

            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Xml, c.Kind));
            Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("root[1]/a:item[1]", StringComparison.Ordinal));
            Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("root[1]/b:item[1]/@b:code", StringComparison.Ordinal));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderText_ReadStructuredTextStream_TrimsSourceNameBeforeDispatch() {
        var json = "{\"service\":{\"name\":\"OfficeIMO\"}}";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(json), writable: false);

        var chunks = DocumentReaderTextExtensions.ReadStructuredText(
            stream,
            sourceName: " config.json ",
            structuredOptions: new StructuredTextReadOptions {
                JsonChunkRows = 10,
                IncludeJsonMarkdown = true
            }).ToList();

        Assert.NotEmpty(chunks);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Json, c.Kind));
        Assert.All(chunks, c => Assert.Equal("config.json", c.Location.Path));
        Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("$.service.name", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderText_CompatibilityRegistration_DispatchesStructuredJsonStream() {
        try {
            DocumentReaderTextRegistrationExtensions.RegisterStructuredTextHandler(replaceExisting: true);

            var payload = "{\"service\":{\"name\":\"IX\",\"enabled\":true,\"ports\":[443,8443]}}";
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(payload), writable: false);
            var chunks = DocumentReader.Read(stream, "config.json").ToList();

            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Json, c.Kind));
            Assert.Contains(chunks, c =>
                (c.Location.Path?.Contains("config.json", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("$.service.name", StringComparison.Ordinal) ?? false));
        } finally {
            DocumentReaderTextRegistrationExtensions.UnregisterStructuredTextHandler();
        }
    }
}

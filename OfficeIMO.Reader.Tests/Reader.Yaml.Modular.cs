using OfficeIMO.Reader;
using OfficeIMO.Reader.Yaml;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderYamlModularTests {
    [Fact]
    public void ReaderInputKind_Yaml_AppendsWithoutChangingExistingAdapterValues() {
        Assert.Equal(10, (int)ReaderInputKind.Html);
        Assert.Equal(11, (int)ReaderInputKind.Zip);
        Assert.Equal(12, (int)ReaderInputKind.Epub);
        Assert.Equal(13, (int)ReaderInputKind.Visio);
        Assert.Equal(14, (int)ReaderInputKind.Yaml);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_ParsesMultiDocumentYamlIntoStructuredChunks() {
        var yaml =
            "apiVersion: v1\n" +
            "kind: Service\n" +
            "metadata:\n" +
            "  name: officeimo\n" +
            "spec:\n" +
            "  ports:\n" +
            "    - name: https\n" +
            "      port: 443\n" +
            "---\n" +
            "enabled: true\n" +
            "replicas: 2\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunks = YamlReaderAdapter.Read(
            stream,
            sourceName: "service.yaml",
            yamlOptions: new YamlReadOptions {
                ChunkRows = 4,
                IncludeMarkdown = true
            }).ToList();

        Assert.NotEmpty(chunks);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Yaml, c.Kind));
        Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("$[0].metadata.name", StringComparison.Ordinal));
        Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("$[0].spec.ports[0].port | number | 443", StringComparison.Ordinal));
        Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("$[1].enabled | boolean | true", StringComparison.Ordinal));
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
    public void DocumentReaderYaml_ReadYamlStream_RespectsExplicitStringTags() {
        const string yaml =
            "flag: !!str true\n" +
            "id: !!str 123\n" +
            "age: !!int \"42\"\n" +
            "enabled: !!bool \"false\"\n" +
            "empty: !!null \"null\"\n" +
            "implicitEnabled: true\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunk = Assert.Single(YamlReaderAdapter.Read(
            stream,
            sourceName: "tagged.yaml",
            yamlOptions: new YamlReadOptions {
                ChunkRows = 10,
                IncludeMarkdown = true
            }));

        Assert.Contains("$.flag | string | true", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("$.id | string | 123", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("$.age | number | 42", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("$.enabled | boolean | false", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("$.empty | null | null", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("$.implicitEnabled | boolean | true", chunk.Text ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_PreservesQuotedScalarWhitespaceInTables() {
        const string yaml = "command: \"printf 'a  b\\n'\"\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunk = Assert.Single(YamlReaderAdapter.Read(
            stream,
            sourceName: "quoted.yaml",
            yamlOptions: new YamlReadOptions {
                ChunkRows = 10,
                IncludeMarkdown = true
            }));

        var table = Assert.Single(chunk.Tables ?? Array.Empty<ReaderTable>());
        var row = Assert.Single(table.Rows);

        Assert.Equal("$.command", row[0]);
        Assert.Equal("string", row[1]);
        Assert.Equal("printf 'a  b\n'", row[2]);
        Assert.Contains("printf 'a  b\\n'", chunk.Text ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_RecognizesYamlCoreNumericForms() {
        const string yaml =
            "mode: 0o755\n" +
            "mask: 0xFF\n" +
            "total: 1_000\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunk = Assert.Single(YamlReaderAdapter.Read(
            stream,
            sourceName: "numbers.yaml",
            yamlOptions: new YamlReadOptions {
                ChunkRows = 10,
                IncludeMarkdown = false
            }));

        Assert.Contains("$.mode | number | 0o755", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("$.mask | number | 0xFF", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("$.total | number | 1_000", chunk.Text ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_EnforcesMaxInputBytesAgainstCompleteSeekableInput() {
        var prefix = Encoding.UTF8.GetBytes(new string('x', 128));
        var yaml = Encoding.UTF8.GetBytes("metadata:\n  name: officeimo\n");
        using var stream = new MemoryStream(prefix.Concat(yaml).ToArray(), writable: false);
        stream.Position = prefix.Length;

        Assert.Throws<IOException>(() => YamlReaderAdapter.Read(
                stream,
                sourceName: "slice.yaml",
                readerOptions: new ReaderOptions { MaxInputBytes = yaml.Length },
                yamlOptions: new YamlReadOptions {
                    ChunkRows = 10,
                    IncludeMarkdown = false
                })
            .ToList());
        Assert.Equal(prefix.Length, stream.Position);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_PreservesBlockScalarValuesInTables() {
        const string yaml =
            "script: |\n" +
            "  echo first\n" +
            "  echo second\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunk = Assert.Single(YamlReaderAdapter.Read(
            stream,
            sourceName: "script.yaml",
            yamlOptions: new YamlReadOptions {
                ChunkRows = 10,
                IncludeMarkdown = true
            }));

        var table = Assert.Single(chunk.Tables ?? Array.Empty<ReaderTable>());
        var row = Assert.Single(table.Rows);

        Assert.Equal("$.script", row[0]);
        Assert.Equal("string", row[1]);
        Assert.Equal("echo first\necho second\n", row[2]);
        Assert.Contains("echo first\\necho second\\n", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("echo first\\\\necho second\\\\n", chunk.Markdown ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_EscapesComplexMappingKeys() {
        const string yaml = "'service.name':\n  'port[0]': 443\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunks = YamlReaderAdapter.Read(
            stream,
            sourceName: "service.yaml",
            yamlOptions: new YamlReadOptions {
                ChunkRows = 10,
                IncludeMarkdown = true
            }).ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("$[\"service.name\"][\"port[0]\"]", StringComparison.Ordinal));
        Assert.Contains(chunks, c => (c.Markdown ?? string.Empty).Contains("$[\"service.name\"][\"port[0]\"]", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_NonSeekable_EnforcesMaxInputBytes() {
        var yaml =
            "metadata:\n" +
            "  name: officeimo\n";
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(yaml));

        var ex = Assert.Throws<IOException>(() => YamlReaderAdapter.Read(
            stream,
            sourceName: "service.yaml",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 },
            yamlOptions: new YamlReadOptions {
                ChunkRows = 2,
                IncludeMarkdown = true
            }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYaml_EmitsWarningForMalformedYaml() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-bad-yaml-" + Guid.NewGuid().ToString("N") + ".yaml");
        try {
            File.WriteAllText(path, "metadata: [unterminated");

            var warningChunk = Assert.Single(
                YamlReaderAdapter.Read(path, new ReaderOptions { ComputeHashes = true }),
                c => c.Kind == ReaderInputKind.Yaml &&
                     (c.Warnings?.Any(w => w.Contains("YAML parse error", StringComparison.OrdinalIgnoreCase)) ?? false));

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
    public void DocumentReaderYaml_BuilderHandler_DispatchesYamlStream() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddYamlHandler().Build();

        var payload = "service:\n  name: IX\n  enabled: true\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(payload), writable: false);
        var chunks = reader.Read(stream, " config.yaml ").ToList();

        Assert.NotEmpty(chunks);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Yaml, c.Kind));
        Assert.All(chunks, c => Assert.Equal("config.yaml", c.Location.Path));
        Assert.Contains(chunks, c => (c.Text?.Contains("$.service.name", StringComparison.Ordinal) ?? false));
        Assert.Equal(ReaderInputKind.Yaml, reader.DetectKind("values.yml"));
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_EnforcesNodeLimitBeforeModelLoad() {
        const string yaml = "root:\n  child:\n    value: 1\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunk = Assert.Single(YamlReaderAdapter.Read(
            stream,
            sourceName: "limited.yaml",
            yamlOptions: new YamlReadOptions {
                MaxNodes = 2,
                ChunkRows = 10,
                IncludeMarkdown = false
            }));

        Assert.Contains("YAML parse limit exceeded", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("maximum node count reached", chunk.Text ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_EnforcesParseEventLimitBeforeModelLoad() {
        const string yaml =
            "root:\n" +
            "  - one\n" +
            "  - two\n" +
            "  - three\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunk = Assert.Single(YamlReaderAdapter.Read(
            stream,
            sourceName: "parse-limited.yaml",
            yamlOptions: new YamlReadOptions {
                MaxParseEvents = 3,
                IncludeMarkdown = false
            }));

        Assert.Contains("YAML parse limit exceeded", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("maximum parse event count reached", chunk.Text ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_EnforcesScalarLengthBeforeModelLoad() {
        string yaml = "value: \"" + new string('a', 32) + "\"\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunk = Assert.Single(YamlReaderAdapter.Read(
            stream,
            sourceName: "scalar-limited.yaml",
            yamlOptions: new YamlReadOptions {
                MaxScalarLength = 8,
                IncludeMarkdown = false
            }));

        Assert.Contains("YAML parse limit exceeded", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("scalar length exceeds maximum", chunk.Text ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_EnforcesDepthLimitBeforeModelLoad() {
        const string yaml =
            "root:\n" +
            "  child:\n" +
            "    grandchild:\n" +
            "      value: one\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunk = Assert.Single(YamlReaderAdapter.Read(
            stream,
            sourceName: "depth-limited.yaml",
            yamlOptions: new YamlReadOptions {
                MaxDepth = 1,
                MaxNodes = 100,
                ChunkRows = 10,
                IncludeMarkdown = false
            }));

        Assert.Contains("YAML parse limit exceeded", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("maximum depth reached", chunk.Text ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_EnforcesScalarDepthBeforeModelLoad() {
        const string yaml =
            "root:\n" +
            "  child:\n" +
            "    value: one\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunk = Assert.Single(YamlReaderAdapter.Read(
            stream,
            sourceName: "scalar-depth-limited.yaml",
            yamlOptions: new YamlReadOptions {
                MaxDepth = 1,
                MaxNodes = 100,
                ChunkRows = 10,
                IncludeMarkdown = false
            }));

        Assert.Contains("YAML parse limit exceeded", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("maximum depth reached", chunk.Text ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderYaml_ReadYamlStream_BoundsComplexMappingKeysWithNodeLimit() {
        const string yaml =
            "? [one, two, three, four]\n" +
            ": value\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(yaml), writable: false);

        var chunk = Assert.Single(YamlReaderAdapter.Read(
            stream,
            sourceName: "complex-key.yaml",
            yamlOptions: new YamlReadOptions {
                MaxNodes = 2,
                ChunkRows = 10,
                IncludeMarkdown = false
            }));

        Assert.Contains("YAML parse limit exceeded", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("maximum node count reached", chunk.Text ?? string.Empty, StringComparison.Ordinal);
    }
}

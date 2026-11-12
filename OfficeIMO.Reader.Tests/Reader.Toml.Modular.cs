using System.Text;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.Toml;
using OfficeIMO.Toml;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderTomlModularTests {
    private const string Pyproject =
        "[project]\n" +
        "name = \"demo\"\n" +
        "version = \"1.0.0\"\n" +
        "dependencies = [\"requests>=2\", \"rich\"]\n" +
        "\n" +
        "[tool.uv.sources]\n" +
        "\"my.pkg\" = { path = \"../pkg\" }\n" +
        "\n" +
        "[[tool.checks]]\n" +
        "when = 1979-05-27\n" +
        "mask = 0xFF\n" +
        "tags = []\n";

    [Fact]
    public void TomlFilesBecomePathTypeValueRowsWhosePathsSelectTheSameValues() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(Pyproject), writable: false);
        var chunks = TomlReaderAdapter.Read(stream, sourceName: "pyproject.toml", tomlOptions: new TomlReadOptions { ChunkRows = 4 }).ToList();

        Assert.Equal(new[] { "toml-0000", "toml-0001" }, chunks.Select(chunk => chunk.Id));
        Assert.All(chunks, chunk => Assert.Equal(ReaderInputKind.Toml, chunk.Kind));
        var rows = chunks.SelectMany(chunk => chunk.Tables![0].Rows).ToList();
        Assert.Equal(new[] {
            "$.project.name|string|demo",
            "$.project.version|string|1.0.0",
            "$.project.dependencies[0]|string|requests>=2",
            "$.project.dependencies[1]|string|rich",
            "$.tool.uv.sources[\"my.pkg\"].path|string|../pkg",
            "$.tool.checks[0].when|date-local|1979-05-27",
            "$.tool.checks[0].mask|integer|0xFF",
            "$.tool.checks[0].tags|array|[]"
        }, rows.Select(row => string.Join("|", row)));
        Assert.Equal((2, 4), (chunks[0].Location.StartLine, chunks[0].Location.EndLine));
        Assert.Contains("| $.project.name | string | demo |", chunks[0].Markdown);

        using TomlDocument document = TomlDocument.Parse(Pyproject);
        foreach (IReadOnlyList<string> row in rows.Where(row => row[1] != "array")) {
            Assert.Equal(row[2], document.Select(row[0])!.AsString());
        }
    }

    [Fact]
    public void InvalidTomlBecomesAWarningChunkWithTheErrorPosition() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("a = 1\na = 2\n"), writable: false);
        ReaderChunk chunk = Assert.Single(TomlReaderAdapter.Read(stream, sourceName: "broken.toml"));
        Assert.Equal("toml-warning-0000", chunk.Id);
        Assert.Contains(TomlErrorCodes.DuplicateKey, chunk.Text);
        Assert.Contains("line 2, column 1", chunk.Text);
    }

    [Fact]
    public void TheAllPresetReadsTomlFilesByExtensionAsTheTomlFormat() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes("title = \"x\"\n"), "config.toml");
        Assert.Equal(ReaderInputKind.Toml, result.Kind);
        Assert.Equal("$.title", Assert.Single(result.Chunks).Tables![0].Rows[0][0]);
        Assert.Equal(OfficeDocumentFormat.Toml, OfficeDocumentReadResultPdfExtensions.MapFormat(result.Kind));
    }

    [Theory]
    [InlineData("MaxDepth", "a.b.c.d.e = 1\n")]
    [InlineData("MaxNodes", "a = 1\nb = 2\nc = 3\n")]
    [InlineData("MaxScalarLength", "a = \"0123456789\"\n")]
    public void TomlLimitsThrowTheNamedReaderLimit(string limit, string toml) {
        var options = new TomlReadOptions {
            MaxDepth = limit == "MaxDepth" ? 2 : null,
            MaxNodes = limit == "MaxNodes" ? 2 : null,
            MaxScalarLength = limit == "MaxScalarLength" ? 8 : null
        };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddTomlHandler(options).Build();

        ReaderResourceLimitException error = Assert.Throws<ReaderResourceLimitException>(() =>
            reader.ReadDocument(Encoding.UTF8.GetBytes(toml), "bounded.toml", new ReaderOptions { ComputeHashes = false }));
        Assert.Equal(limit, error.LimitName);
        Assert.IsType<TomlLimitExceededException>(error.InnerException);
    }

    [Theory]
    [InlineData(null, true)]
    [InlineData(100, false)]
    public void ReaderLimitsOverrideTheNestedLoadOptions(int? maxNodes, bool exceeds) {
        var options = new TomlReadOptions { LoadOptions = new TomlLoadOptions { MaxNodes = 2 }, MaxNodes = maxNodes };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddTomlHandler(options).Build();
        Func<OfficeDocumentReadResult> read = () => reader.ReadDocument(Encoding.UTF8.GetBytes("a = 1\nb = 2\nc = 3\n"), "nested.toml");

        if (exceeds) {
            Assert.Equal("MaxNodes", Assert.Throws<ReaderResourceLimitException>(() => read()).LimitName);
        } else {
            Assert.Equal(3, Assert.Single(read().Chunks).Tables![0].Rows.Count);
        }
    }

    [Fact]
    public void TomlReadStopsWhenCancelled() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(Pyproject), writable: false);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() => TomlReaderAdapter.Read(stream, "pyproject.toml", cancellationToken: cancellation.Token).ToList());
    }
}

using System.Text;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class QueryCommandTests {
    private static readonly (string Extension, string Content)[] Equivalent = [
        (".yaml", "services:\n  - name: web\n    port: 8080\n    tags: [a, b]\n  - name: api\n    port: \"8080\"\n"),
        (".json", "{\"services\": [{\"name\": \"web\", \"port\": 8080, \"tags\": [\"a\", \"b\"]}, {\"name\": \"api\", \"port\": \"8080\"}]}"),
        (".toml", "[[services]]\nname = \"web\"\nport = 8080\ntags = [\"a\", \"b\"]\n\n[[services]]\nname = \"api\"\nport = \"8080\"\n")
    ];

    [Theory]
    [InlineData("services[?(@.port == 8080)].tags[-1]", "services[0].tags[1]|string|b")]
    [InlineData("..name", "services[0].name|string|web;services[1].name|string|api")]
    [InlineData("services[*].port", "services[0].port|number|8080;services[1].port|string|8080")]
    [InlineData("services[?(@.name == 'web')]", "services[0]|mapping|{\"name\":\"web\",\"port\":8080,\"tags\":[\"a\",\"b\"]}")]
    public async Task OneQueryReadsYamlJsonAndTomlAlike(string query, string expected) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-query-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            foreach ((string extension, string content) in Equivalent) {
                string file = Path.Combine(directory, "services" + extension);
                File.WriteAllText(file, content);
                using var output = new MemoryStream();
                using var error = new StringWriter();

                int exitCode = await OfficeImoToolApp.RunAsync(["query", file, query, "--format", "json"], Stream.Null, output, error);

                Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
                Assert.Equal(string.Empty, error.ToString());
                using JsonDocument json = JsonDocument.Parse(output.ToArray());
                string actual = string.Join(";", json.RootElement.EnumerateArray().Select(match =>
                    match.GetProperty("path").GetString() + "|" + match.GetProperty("kind").GetString() + "|" +
                    (match.GetProperty("value").ValueKind == JsonValueKind.String ? match.GetProperty("value").GetString() : JsonSerializer.Serialize(match.GetProperty("value")))));
                Assert.True(expected == actual, extension + ": " + actual);
                Assert.All(json.RootElement.EnumerateArray(), match => Assert.True(match.GetProperty("line").GetInt32() > 0, extension));
            }
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task TextOutputPrintsOneLinePerMatchWithItsPosition() {
        string file = Path.Combine(Path.GetTempPath(), "officeimo-query-" + Guid.NewGuid().ToString("N") + ".yaml");
        File.WriteAllText(file, "a:\n  note: \"two\\nlines\"\n  list: [1, 2]\n");
        try {
            using var output = new MemoryStream();
            using var error = new StringWriter();

            int exitCode = await OfficeImoToolApp.RunAsync(["query", file, "a.*"], Stream.Null, output, error);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
            Assert.Equal("a.note\t2:9\ttwo\\nlines\na.list\t3:9\t[1,2]\n", Encoding.UTF8.GetString(output.ToArray()).Replace("\r\n", "\n"));
        } finally {
            File.Delete(file);
        }
    }

    [Theory]
    [InlineData("services[", "malformed at position 9")]
    [InlineData("--format", "--format needs a value")]
    public async Task BadQueriesAndOptionsAreUsageErrors(string argument, string message) {
        string file = Path.Combine(Path.GetTempPath(), "officeimo-query-" + Guid.NewGuid().ToString("N") + ".toml");
        File.WriteAllText(file, "x = 1\n");
        try {
            using var output = new MemoryStream();
            using var error = new StringWriter();

            int exitCode = await OfficeImoToolApp.RunAsync(["query", file, argument], Stream.Null, output, error);

            Assert.Equal((int)OfficeImoToolExitCode.Usage, exitCode);
            Assert.Contains(message, error.ToString(), StringComparison.Ordinal);
            Assert.Equal(0, output.Length);
        } finally {
            File.Delete(file);
        }
    }

    [Fact]
    public async Task UnknownExtensionsNeedAnInputFormat() {
        string file = Path.Combine(Path.GetTempPath(), "officeimo-query-" + Guid.NewGuid().ToString("N") + ".conf");
        File.WriteAllText(file, "x = 1\n");
        try {
            using var output = new MemoryStream();
            using var error = new StringWriter();

            Assert.Equal((int)OfficeImoToolExitCode.Usage, await OfficeImoToolApp.RunAsync(["query", file, "x"], Stream.Null, output, error));
            Assert.Equal((int)OfficeImoToolExitCode.Success, await OfficeImoToolApp.RunAsync(["query", file, "x", "--input-format", "toml"], Stream.Null, output, error));
            Assert.Equal("x\t1:5\t1\n", Encoding.UTF8.GetString(output.ToArray()).Replace("\r\n", "\n"));
        } finally {
            File.Delete(file);
        }
    }
}

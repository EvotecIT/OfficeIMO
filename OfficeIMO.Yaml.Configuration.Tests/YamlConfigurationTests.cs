using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Options;
using OfficeIMO.Yaml;
using Xunit;

namespace OfficeIMO.Yaml.Configuration.Tests;

public sealed class YamlConfigurationTests {
    [Fact]
    public void BindsNestedObjectsListsAndDictionariesThroughOptions() {
        IConfiguration configuration = Build("""
            App:
              Name: Portal
              Port: 8443
              Enabled: TRUE
              Started: 2024-05-01T08:30:00Z
              Database:
                Host: db.local
                Timeout: 00:00:30
              Hosts:
                - a.example
                - b.example
              Endpoints:
                - Name: primary
                  Weight: 1
                - Name: backup
                  Weight: 2
              Limits:
                requests: 100
                Uploads: 5
            """);

        var services = new ServiceCollection();
        services.Configure<AppOptions>(configuration.GetSection("App"));
        using ServiceProvider provider = services.BuildServiceProvider();
        AppOptions options = provider.GetRequiredService<IOptions<AppOptions>>().Value;

        Assert.Equal("Portal", options.Name);
        Assert.Equal(8443, options.Port);
        Assert.True(options.Enabled);
        Assert.Equal(new DateTimeOffset(2024, 5, 1, 8, 30, 0, TimeSpan.Zero), options.Started);
        Assert.Equal("db.local", options.Database.Host);
        Assert.Equal(TimeSpan.FromSeconds(30), options.Database.Timeout);
        Assert.Equal(new[] { "a.example", "b.example" }, options.Hosts);
        Assert.Equal(new[] { "primary", "backup" }, options.Endpoints.Select(e => e.Name));
        Assert.Equal(new[] { 1, 2 }, options.Endpoints.Select(e => e.Weight));
        Assert.Equal(100, options.Limits["requests"]);
        Assert.Equal(5, options.Limits["uploads"]);
    }

    [Fact]
    public void FlattensKeysAsTheJsonProviderDoes() {
        IDictionary<string, string?> data = Load("""
            Logging:
              LogLevel:
                Default: Information
            Hosts: [a, b]
            Empty: []
            Nothing: {}
            Blank:
            Tilde: ~
            Text: "null"
            Flag: TRUE
            Hex: 0x1F
            Block: |
              line one
              line two
            After: x
            """);

        Assert.Equal("Information", data["Logging:LogLevel:Default"]);
        Assert.Equal("a", data["Hosts:0"]);
        Assert.Equal("b", data["Hosts:1"]);
        Assert.Null(data["Empty"]);
        Assert.Null(data["Nothing"]);
        Assert.Equal(string.Empty, data["Blank"]);
        Assert.Equal(string.Empty, data["Tilde"]);
        Assert.Equal("null", data["Text"]);
        Assert.Equal("TRUE", data["Flag"]);
        Assert.Equal("0x1F", data["Hex"]);
        Assert.Equal("line one\nline two\n", data["Block"]);
        Assert.Equal("Information", data["logging:loglevel:default"]);
    }

    [Fact]
    public void ResolvesAliasesAndMergeKeys() {
        IDictionary<string, string?> data = Load("""
            defaults: &defaults
              Timeout: 30
              Retries: 3
              Tags: [base]
            extra: &extra
              Retries: 9
              Region: west
            Primary:
              <<: [*defaults, *extra]
              Timeout: 10
            Secondary:
              <<: *defaults
              Copy: *extra
            """);

        Assert.Equal("10", data["Primary:Timeout"]);
        Assert.Equal("3", data["Primary:Retries"]);
        Assert.Equal("west", data["Primary:Region"]);
        Assert.Equal("base", data["Primary:Tags:0"]);
        Assert.False(data.ContainsKey("Primary:<<"));
        Assert.Equal("30", data["Secondary:Timeout"]);
        Assert.Equal("9", data["Secondary:Copy:Retries"]);
    }

    [Fact]
    public void ResolvesMergeKeysUnderTheYaml11Schema() {
        IDictionary<string, string?> data = Load("""
            %YAML 1.1
            ---
            base: &base
              Enabled: yes
            Service:
              <<: *base
              Name: api
            """);

        Assert.Equal("yes", data["Service:Enabled"]);
        Assert.Equal("api", data["Service:Name"]);
    }

    [Fact]
    public void KeepsAnEmptyKeyApartFromTheRoot() {
        IDictionary<string, string?> data = Load("\"\":\n  Name: nested\nName: root\nOuter:\n  \"\": {}\n");

        Assert.Equal("nested", data[":Name"]);
        Assert.Equal("root", data["Name"]);
        Assert.True(data.ContainsKey("Outer:"));
        Assert.Null(data["Outer:"]);
    }

    [Fact]
    public void RejectsMoreThanOneDocument() {
        FormatException exception = Assert.Throws<FormatException>(() => Load("""
            Name: first
            ---
            Name: second
            """));

        Assert.Contains("one document", exception.Message);
        Assert.Contains("line 3", exception.Message);
    }

    [Fact]
    public void ReadsOneDocumentWithMarkers() {
        IDictionary<string, string?> data = Load("""
            ---
            Name: only
            ...
            """);

        Assert.Equal("only", data["Name"]);
    }

    [Theory]
    [InlineData("")]
    [InlineData("# only a comment\n")]
    [InlineData("~\n")]
    public void ReadsEmptyDocumentsAsNoKeys(string yaml) {
        Assert.Empty(Load(yaml));
    }

    [Fact]
    public void RejectsARootThatIsNotAMapping() {
        FormatException exception = Assert.Throws<FormatException>(() => Load("- a\n- b\n"));
        Assert.Contains("must be a mapping", exception.Message);
    }

    [Theory]
    [InlineData("Server:\n  Port: 1\n  Port: 2\n", "Server:Port", 3)]
    [InlineData("Server:\n  Port: 1\n  port: 2\n", "Server:port", 3)]
    [InlineData("a:b: 1\na:\n  b: 2\n", "a:b", 3)]
    [InlineData("\"a:b\": 1\na:\n  b: []\n", "a:b", 3)]
    [InlineData("a:\n  b: {}\n\"a:b\": 1\n", "a:b", 3)]
    public void RejectsDuplicateKeysWithTheirLine(string yaml, string key, int line) {
        FormatException exception = Assert.Throws<FormatException>(() => Load(yaml));
        Assert.Equal($"A duplicate key '{key}' was found at line {line}.", exception.Message);
    }

    [Fact]
    public void SurfacesParseErrorsWithTheirPosition() {
        string path = WriteTempFile("app.yaml", "Server:\n  Host: local\n Port: [1, 2\n");
        try {
            var builder = new ConfigurationBuilder().AddYamlFile(path);
            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => builder.Build());

            FormatException format = Assert.IsType<FormatException>(exception.InnerException);
            YamlException yaml = Assert.IsType<YamlException>(format.InnerException);
            Assert.Equal(3, yaml.Line);
            Assert.Contains("line 3", format.Message);
        } finally {
            Directory.Delete(Path.GetDirectoryName(path)!, recursive: true);
        }
    }

    [Fact]
    public void ToleratesAMissingOptionalFile() {
        string directory = CreateTempDirectory();
        try {
            IConfiguration configuration = new ConfigurationBuilder()
                .SetBasePath(directory)
                .AddYamlFile("missing.yaml", optional: true)
                .Build();

            Assert.Empty(configuration.AsEnumerable());
            Assert.Throws<FileNotFoundException>(() => new ConfigurationBuilder().SetBasePath(directory).AddYamlFile("missing.yaml").Build());
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ReloadsWhenTheFileChanges() {
        string path = WriteTempFile("app.yaml", "Feature:\n  Enabled: false\n");
        string directory = Path.GetDirectoryName(path)!;
        try {
            using var files = new PhysicalFileProvider(directory) { UsePollingFileWatcher = true, UseActivePolling = true };
            IConfigurationRoot configuration = new ConfigurationBuilder()
                .AddYamlFile(source => {
                    source.FileProvider = files;
                    source.Path = "app.yaml";
                    source.ReloadOnChange = true;
                    source.ReloadDelay = 0;
                })
                .Build();
            Assert.Equal("false", configuration["Feature:Enabled"]);

            using var reloaded = new ManualResetEventSlim();
            using IDisposable registration = Microsoft.Extensions.Primitives.ChangeToken.OnChange(configuration.GetReloadToken, reloaded.Set);
            File.WriteAllText(path, "Feature:\n  Enabled: true\n  Name: new\n");
            File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddSeconds(5));

            Assert.True(reloaded.Wait(TimeSpan.FromSeconds(30)), "The configuration did not reload.");
            Assert.True(WaitFor(() => configuration["Feature:Name"] == "new"));
            Assert.Equal("true", configuration["Feature:Enabled"]);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ReadsAStream() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Name: streamed\nItems: [1, 2]\n"));
        IConfiguration configuration = new ConfigurationBuilder().AddYamlStream(stream).Build();

        Assert.Equal("streamed", configuration["Name"]);
        Assert.Equal("2", configuration["Items:1"]);
    }

    [Fact]
    public void MatchesTheJsonProvider() {
        const string json = """
            {
              "Logging": { "LogLevel": { "Default": "Information", "Microsoft.AspNetCore": "Warning" } },
              "Server": {
                "Host": "localhost",
                "Port": 8080,
                "Enabled": true,
                "Ratio": 0.75,
                "Proxy": null,
                "Tags": [ "web", "api" ],
                "Empty": [],
                "Nothing": {},
                "Matrix": [ [ 1, 2 ], [ 3 ] ]
              },
              "Endpoints": [
                { "Name": "primary", "Url": "https://a.example" },
                { "Name": "backup", "Url": "https://b.example", "Weight": 2 }
              ],
              "ConnectionStrings": { "Default": "Server=.;Database=app" }
            }
            """;
        const string yaml = """
            Logging:
              LogLevel:
                Default: Information
                Microsoft.AspNetCore: Warning
            Server:
              Host: localhost
              Port: 8080
              Enabled: true
              Ratio: 0.75
              Proxy: null
              Tags: [web, api]
              Empty: []
              Nothing: {}
              Matrix:
                - [1, 2]
                - [3]
            Endpoints:
              - Name: primary
                Url: https://a.example
              - Name: backup
                Url: https://b.example
                Weight: 2
            ConnectionStrings:
              Default: Server=.;Database=app
            """;

        IConfiguration fromJson = new ConfigurationBuilder().AddJsonStream(Utf8(json)).Build();
        IConfiguration fromYaml = new ConfigurationBuilder().AddYamlStream(Utf8(yaml)).Build();

        ConfigurationParity.AssertEquivalent(fromJson, fromYaml);
    }

    private static IConfiguration Build(string yaml) => new ConfigurationBuilder().AddYamlStream(Utf8(yaml)).Build();

    private static IDictionary<string, string?> Load(string yaml) {
        var provider = new ProbeProvider(Utf8(yaml));
        provider.Load();
        return provider.Values;
    }

    private sealed class ProbeProvider : YamlStreamConfigurationProvider {
        public ProbeProvider(Stream stream) : base(new YamlStreamConfigurationSource { Stream = stream }) { }

        public IDictionary<string, string?> Values => Data;
    }

    private static MemoryStream Utf8(string text) => new MemoryStream(Encoding.UTF8.GetBytes(text));

    private static bool WaitFor(Func<bool> condition) {
        DateTime deadline = DateTime.UtcNow.AddSeconds(10);
        while (DateTime.UtcNow < deadline) {
            if (condition()) return true;
            Thread.Sleep(50);
        }
        return condition();
    }

    private static string CreateTempDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-yaml-config-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static string WriteTempFile(string name, string content) {
        string path = Path.Combine(CreateTempDirectory(), name);
        File.WriteAllText(path, content);
        return path;
    }
}

using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Options;
using Microsoft.Extensions.Primitives;
using OfficeIMO.Toml;
using Xunit;

namespace OfficeIMO.Toml.Configuration.Tests;

public sealed class TomlConfigurationTests {
    [Fact]
    public void BindsNestedObjectsListsAndDictionariesThroughOptions() {
        IConfiguration configuration = Build("""
            [App]
            Name = "Portal"
            Port = 8443
            Enabled = true
            Started = 2024-05-01 08:30:00Z
            Hosts = ["a.example", "b.example"]
            Database = { Host = "db.local", Timeout = "00:00:30" }
            Limits.requests = 100
            Limits.Uploads = 5

            [[App.Endpoints]]
            Name = "primary"
            Weight = 1

            [[App.Endpoints]]
            Name = "backup"
            Weight = 2
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
    public void FlattensTablesDottedKeysAndInlineTablesAlike() {
        IDictionary<string, string?> data = Load("""
            title = "app"
            server.host = "localhost"
            client = { retries = 3, backoff = { initial = 1 } }

            [database]
            ports = [8000, 8001]
            matrix = [[1, 2], [3]]
            empty = []
            nothing = {}

            [database.replica]
            host = "replica"

            [[products]]
            name = "Hammer"

            [[products]]

            [[products]]
            name = "Nail"
            """);

        Assert.Equal("app", data["title"]);
        Assert.Equal("localhost", data["server:host"]);
        Assert.Equal("3", data["client:retries"]);
        Assert.Equal("1", data["client:backoff:initial"]);
        Assert.Equal("8001", data["database:ports:1"]);
        Assert.Equal("3", data["database:matrix:1:0"]);
        Assert.Null(data["database:empty"]);
        Assert.Null(data["database:nothing"]);
        Assert.Equal("replica", data["database:replica:host"]);
        Assert.Equal("Hammer", data["products:0:name"]);
        Assert.Null(data["products:1"]);
        Assert.Equal("Nail", data["products:2:name"]);
        Assert.Equal("Nail", data["PRODUCTS:2:NAME"]);
    }

    [Fact]
    public void WritesValuesSoTheyBind() {
        IDictionary<string, string?> data = Load("""
            text = 'C:\path'
            multiline = '''
            one
            two'''
            decimal = 1500
            negative = -17
            grouped = 1_000
            hex = 0xFF
            octal = 0o17
            binary = 0b101
            positive = +99
            float = 3.0
            exponent = 6.626e-34
            groupedFloat = 224_617.445_991
            infinity = inf
            negativeInfinity = -inf
            notANumber = nan
            flag = true
            offset = 1979-05-27 07:32:00.5-07:00
            local = 1979-05-27T07:32:00
            date = 1979-05-27
            time = 07:32:00
            """);

        Assert.Equal("C:\\path", data["text"]);
        Assert.Equal("one\ntwo", data["multiline"]);
        Assert.Equal("1500", data["decimal"]);
        Assert.Equal("-17", data["negative"]);
        Assert.Equal("1000", data["grouped"]);
        Assert.Equal("255", data["hex"]);
        Assert.Equal("15", data["octal"]);
        Assert.Equal("5", data["binary"]);
        Assert.Equal("99", data["positive"]);
        Assert.Equal("3.0", data["float"]);
        Assert.Equal("6.626e-34", data["exponent"]);
        Assert.Equal("224617.445991", data["groupedFloat"]);
        Assert.Equal("Infinity", data["infinity"]);
        Assert.Equal("-Infinity", data["negativeInfinity"]);
        Assert.Equal("NaN", data["notANumber"]);
        Assert.Equal("true", data["flag"]);
        Assert.Equal("1979-05-27T07:32:00.5-07:00", data["offset"]);
        Assert.Equal("1979-05-27T07:32:00", data["local"]);
        Assert.Equal("1979-05-27", data["date"]);
        Assert.Equal("07:32:00", data["time"]);

        IConfiguration configuration = new ConfigurationBuilder().AddInMemoryCollection(data).Build();
        Assert.Equal(255, configuration.GetValue<int>("hex"));
        Assert.Equal(double.PositiveInfinity, configuration.GetValue<double>("infinity"));
        Assert.True(double.IsNaN(configuration.GetValue<double>("notANumber")));
        Assert.Equal(new DateTimeOffset(1979, 5, 27, 7, 32, 0, 500, TimeSpan.FromHours(-7)), configuration.GetValue<DateTimeOffset>("offset"));
        Assert.Equal(new DateTime(1979, 5, 27, 7, 32, 0), configuration.GetValue<DateTime>("local"));
        Assert.Equal(new DateTime(1979, 5, 27), configuration.GetValue<DateTime>("date"));
        Assert.Equal(new TimeSpan(7, 32, 0), configuration.GetValue<TimeSpan>("time"));
    }

    [Theory]
    [InlineData("[server]\nPort = 1\nport = 2\n", "server:port", 3)]
    [InlineData("Name = \"a\"\nname.first = \"b\"\n", "name", 2)]
    [InlineData("[A]\nx = 1\n[a]\ny = 2\n", "a", 3)]
    [InlineData("\"a:b\" = 1\na = { b = [] }\n", "a:b", 2)]
    [InlineData("a = { b = {} }\n\"a:b\" = 1\n", "a:b", 2)]
    public void RejectsKeysThatDifferOnlyInCase(string toml, string key, int line) {
        FormatException exception = Assert.Throws<FormatException>(() => Load(toml));
        Assert.Equal($"A duplicate key '{key}' was found at line {line}.", exception.Message);
    }

    [Fact]
    public void KeepsAnEmptyKeyApartFromTheRoot() {
        IDictionary<string, string?> data = Load("Port = 25\n\"\" = { Port = 587 }\nOuter = { \"\" = {} }\n");

        Assert.Equal("587", data[":Port"]);
        Assert.Equal("25", data["Port"]);
        Assert.True(data.ContainsKey("Outer:"));
        Assert.Null(data["Outer:"]);
    }

    [Fact]
    public void RejectsDuplicateKeysWithTheirPosition() {
        FormatException exception = Assert.Throws<FormatException>(() => Load("[server]\nport = 1\nport = 2\n"));

        TomlException toml = Assert.IsType<TomlException>(exception.InnerException);
        Assert.Equal(TomlErrorCodes.DuplicateKey, toml.Code);
        Assert.Equal(3, toml.Line);
        Assert.Contains("line 3", exception.Message);
        Assert.Contains("port", exception.Message);
    }

    [Fact]
    public void SurfacesParseErrorsWithTheirPosition() {
        string path = WriteTempFile("app.toml", "[server]\nhost = \"local\"\nport = [1, 2\n");
        try {
            var builder = new ConfigurationBuilder().AddTomlFile(path);
            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => builder.Build());

            FormatException format = Assert.IsType<FormatException>(exception.InnerException);
            TomlException toml = Assert.IsType<TomlException>(format.InnerException);
            Assert.True(toml.Line >= 3);
            Assert.Contains("line " + toml.Line.ToString(System.Globalization.CultureInfo.InvariantCulture), format.Message);
        } finally {
            Directory.Delete(Path.GetDirectoryName(path)!, recursive: true);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("# only a comment\n")]
    public void ReadsEmptyDocumentsAsNoKeys(string toml) {
        Assert.Empty(Load(toml));
    }

    [Fact]
    public void ToleratesAMissingOptionalFile() {
        string directory = CreateTempDirectory();
        try {
            IConfiguration configuration = new ConfigurationBuilder()
                .SetBasePath(directory)
                .AddTomlFile("missing.toml", optional: true)
                .Build();

            Assert.Empty(configuration.AsEnumerable());
            Assert.Throws<FileNotFoundException>(() => new ConfigurationBuilder().SetBasePath(directory).AddTomlFile("missing.toml").Build());
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ReloadsWhenTheFileChanges() {
        string path = WriteTempFile("app.toml", "[Feature]\nEnabled = false\n");
        string directory = Path.GetDirectoryName(path)!;
        try {
            using var files = new PhysicalFileProvider(directory) { UsePollingFileWatcher = true, UseActivePolling = true };
            IConfigurationRoot configuration = new ConfigurationBuilder()
                .AddTomlFile(files, "app.toml", optional: false, reloadOnChange: true)
                .Build();
            Assert.Equal("false", configuration["Feature:Enabled"]);

            using var reloaded = new ManualResetEventSlim();
            using IDisposable registration = ChangeToken.OnChange(configuration.GetReloadToken, reloaded.Set);
            File.WriteAllText(path, "[Feature]\nEnabled = true\nName = \"new\"\n");
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
        using MemoryStream stream = Utf8("Name = \"streamed\"\nItems = [1, 2]\n");
        IConfiguration configuration = new ConfigurationBuilder().AddTomlStream(stream).Build();

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
        const string toml = """
            [Logging.LogLevel]
            Default = "Information"
            "Microsoft.AspNetCore" = "Warning"

            [Server]
            Host = "localhost"
            Port = 8080
            Enabled = true
            Ratio = 0.75
            Tags = ["web", "api"]
            Empty = []
            Nothing = {}
            Matrix = [[1, 2], [3]]

            [[Endpoints]]
            Name = "primary"
            Url = "https://a.example"

            [[Endpoints]]
            Name = "backup"
            Url = "https://b.example"
            Weight = 2

            [ConnectionStrings]
            Default = "Server=.;Database=app"
            """;

        IConfiguration fromJson = new ConfigurationBuilder().AddJsonStream(Utf8(json)).Build();
        IConfiguration fromToml = new ConfigurationBuilder().AddTomlStream(Utf8(toml)).Build();

        ConfigurationParity.AssertEquivalent(fromJson, fromToml);
    }

    private static IConfiguration Build(string toml) => new ConfigurationBuilder().AddTomlStream(Utf8(toml)).Build();

    private static IDictionary<string, string?> Load(string toml) {
        var provider = new ProbeProvider(Utf8(toml));
        provider.Load();
        return provider.Values;
    }

    private sealed class ProbeProvider : TomlStreamConfigurationProvider {
        public ProbeProvider(Stream stream) : base(new TomlStreamConfigurationSource { Stream = stream }) { }

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
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-toml-config-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static string WriteTempFile(string name, string content) {
        string path = Path.Combine(CreateTempDirectory(), name);
        File.WriteAllText(path, content);
        return path;
    }
}

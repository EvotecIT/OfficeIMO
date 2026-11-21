using System.Text.Json;
using OfficeIMO.CSV;
using OfficeIMO.Data.Conversions;
using OfficeIMO.Toml;
using Xunit;

namespace OfficeIMO.Data.Conversions.Tests;

// TOML sink layout and comment carrying, and the TSV dialect of the table routes.
public class TomlTsvConversionTests {
    // Values first, then [table] and [[array]] sections; arrays of scalars and mixed arrays stay inline.
    [Fact]
    public void TomlOutputUsesSectionsForObjectsAndArraysOfObjects() {
        const string json = """
            {"name": "app", "server": {"host": "localhost", "ports": [80, 443]},
             "package": [{"name": "a", "deps": {"x": "1"}}, {"name": "b"}],
             "mixed": [1, {"k": "v"}], "empty": {}, "version": "1.0"}
            """;
        (string toml, OfficeConversionReport report) = Convert.Run(json, DataFormat.Json, DataFormat.Toml);
        Assert.Equal("""
            name = "app"
            mixed = [1, { k = "v" }]
            version = "1.0"

            [server]
            host = "localhost"
            ports = [80, 443]

            [[package]]
            name = "a"

            [package.deps]
            x = "1"

            [[package]]
            name = "b"

            [empty]

            """.Replace("\r\n", "\n"), toml);
        Assert.Contains(OfficeDiagnosticCodes.KeyOrderChanged, Convert.Codes(report));
        (string back, _) = Convert.Run(toml, DataFormat.Toml, DataFormat.Json);
        using JsonDocument data = JsonDocument.Parse(back);
        Assert.Equal("a", data.RootElement.GetProperty("package")[0].GetProperty("name").GetString());
        Assert.Equal("1", data.RootElement.GetProperty("package")[0].GetProperty("deps").GetProperty("x").GetString());
    }

    [Fact]
    public void TomlCommentsTravelThroughYaml() {
        const string toml = "# the title\ntitle = \"x\" # inline\n\n# owner section\n[owner]\nname = \"Tom\"\n";
        (string yaml, OfficeConversionReport toYaml) = Convert.Run(toml, DataFormat.Toml, DataFormat.Yaml);
        Assert.Contains("the title", yaml);
        Assert.Contains("owner section", yaml);
        Assert.DoesNotContain(OfficeDiagnosticCodes.CommentsOmitted, Convert.Codes(toYaml));
        (string back, _) = Convert.Run(yaml, DataFormat.Yaml, DataFormat.Toml);
        using TomlDocument document = TomlDocument.Parse(back);
        Assert.Contains(document.Root["title"].LeadingComments, c => c.Contains("the title"));
        Assert.Contains(document.Root["owner"].LeadingComments, c => c.Contains("owner section"));
    }

    [Fact]
    public void TomlDateTimeKindsSurviveTablesAndTomlAgain() {
        const string toml = "[[event]]\nat = 1979-05-27T07:32:00.123-07:00\nlocal = 1979-05-27T07:32:00\nday = 1979-05-27\ntime = 07:32:00\n";
        (string csv, _) = Convert.Run(toml, DataFormat.Toml, DataFormat.Csv, new DataConversionOptions { Select = "event" });
        Assert.Equal("at,local,day,time\n1979-05-27T07:32:00.123-07:00,1979-05-27T07:32:00,1979-05-27,07:32:00\n", csv);
        (string back, _) = Convert.Run(csv, DataFormat.Csv, DataFormat.Toml);
        using TomlDocument document = TomlDocument.Parse(back);
        TomlTable row = document.Root["root"].AsTableArray()[0];
        Assert.Equal(TomlDateTimeKind.OffsetDateTime, row["at"].GetValue<TomlDateTime>().Kind);
        Assert.Equal(TomlDateTimeKind.LocalDateTime, row["local"].GetValue<TomlDateTime>().Kind);
        Assert.Equal(TomlDateTimeKind.LocalDate, row["day"].GetValue<TomlDateTime>().Kind);
        // CSV inference types dates and date-times, not times of day, so a local time comes back as text.
        Assert.Equal("07:32:00", row["time"].AsString());
    }

    // TSV is tab-separated with no quoting: quotes are data, and a value with a tab or line break fails instead of
    // producing a file that reads back differently.
    [Fact]
    public void TsvRoutesUseTheTabProfile() {
        const string rows = "[{\"id\":1,\"note\":\"say \\\"hi\\\", twice\",\"zip\":\"007\"},{\"id\":2,\"note\":\"plain\",\"zip\":\"120\"}]";
        (string tsv, _) = Convert.Run(rows, DataFormat.Json, DataFormat.Tsv);
        Assert.Equal("id\tnote\tzip\n1\tsay \"hi\", twice\t007\n2\tplain\t120\n", tsv);
        (string json, _) = Convert.Run(tsv, DataFormat.Tsv, DataFormat.Json);
        using JsonDocument back = JsonDocument.Parse(json);
        Assert.Equal("say \"hi\", twice", back.RootElement[0].GetProperty("note").GetString());
        Assert.Equal(1, back.RootElement[0].GetProperty("id").GetInt32());
        Assert.Equal("007", back.RootElement[0].GetProperty("zip").GetString());

        Assert.ThrowsAny<Exception>(() => Convert.Run("[{\"a\":\"tab\\there\"}]", DataFormat.Json, DataFormat.Tsv));
        var escaped = new DataConversionOptions { CsvSave = CsvProfiles.CreateSaveOptions(CsvProfile.Tsv) };
        escaped.CsvSave.EscapeMode = CsvEscapeMode.Backslash;
        (string withEscapes, _) = Convert.Run("[{\"a\":\"tab\\there\"}]", DataFormat.Json, DataFormat.Tsv, escaped);
        Assert.Equal("a\ntab\\there\n", withEscapes);
    }

    [Fact]
    public void ModelSaveAsCsvWritesTsvForTsvPaths() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-data-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            using TomlDocument cargo = TomlDocument.Parse("[[package]]\nname = \"serde\"\nversion = \"1.0.200\"\n[[package]]\nname = \"toml\"\nversion = \"0.8.0\"\n");
            cargo.SaveAsCsv(Path.Combine(directory, "lock.tsv"), new DataConversionOptions { Select = "package" });
            Assert.Equal("name\tversion\nserde\t1.0.200\ntoml\t0.8.0\n", File.ReadAllText(Path.Combine(directory, "lock.tsv")));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}

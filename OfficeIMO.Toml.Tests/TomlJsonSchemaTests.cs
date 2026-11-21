using OfficeIMO.Json;
using Xunit;

namespace OfficeIMO.Toml.Tests;

// TOML documents validate as their JSON data: tables are objects, arrays of tables are arrays, date-times are RFC 3339
// strings, and findings carry the TOML line and column.
public class TomlJsonSchemaTests {
    [Fact]
    public void TomlValuesValidateWithTheirPositions() {
        JsonSchema schema = JsonSchema.Parse("""
            {
              "type": "object",
              "required": ["package"],
              "properties": {
                "package": {
                  "type": "object",
                  "required": ["name", "version"],
                  "properties": {
                    "name": { "type": "string" },
                    "edition": { "enum": ["2018", "2021"] },
                    "published": { "type": "string", "format": "date-time" }
                  }
                },
                "bin": { "type": "array", "items": { "required": ["name", "path"] } }
              }
            }
            """);
        TomlDocument document = TomlDocument.Parse("""
            [package]
            name = "demo"
            edition = 2021
            published = 1979-05-27T07:32:00Z

            [[bin]]
            name = "cli"
            path = "src/cli.rs"

            [[bin]]
            name = "tool"
            """);

        JsonSchemaValidationResult result = document.Validate(schema);

        Assert.Collection(result.Findings,
            version => {
                Assert.Equal("/package", version.InstancePath);
                Assert.Equal("Required property 'version' is missing.", version.Message);
            },
            edition => {
                Assert.Equal("/package/edition", edition.InstancePath);
                Assert.Equal((3, 11), (edition.Line, edition.Column));
                Assert.Same(document.Root.Select("package.edition"), edition.Source);
            },
            path => {
                Assert.Equal("/bin/1", path.InstancePath);
                Assert.Equal("required", path.Keyword);
            });
    }

    [Fact]
    public void OnlyOffsetDateTimesAndLocalDatesMatchDateFormats() {
        JsonSchema schema = JsonSchema.Parse("""
            {
              "properties": {
                "offset": { "format": "date-time" },
                "day": { "format": "date" },
                "local": { "format": "date-time" },
                "clock": { "format": "time" }
              }
            }
            """);
        TomlDocument document = TomlDocument.Parse("""
            offset = 1979-05-27T07:32:00Z
            day = 1979-05-27
            local = 1979-05-27T07:32:00
            clock = 07:32:00
            """);

        Assert.Equal(new[] { "/local", "/clock" }, document.Validate(schema).Findings.Select(f => f.InstancePath));
    }

    [Fact]
    public void DeeplyNestedEditedArraysFailToConvert() {
        TomlDocument document = TomlDocument.Create();
        TomlArray array = document.Root.AddArray("deep");
        for (int i = 0; i < 600; i++) array = array.AddArray();

        Assert.Throws<InvalidOperationException>(() => document.ToJsonData());
    }
}

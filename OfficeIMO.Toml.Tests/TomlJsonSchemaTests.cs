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
}

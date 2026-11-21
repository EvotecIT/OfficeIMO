using OfficeIMO.Json;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// What the official suite cannot see: YAML documents validate with the types their schema resolves, findings point
// at the failing value's line, column and JSON Pointer, and untrusted schemas stay within their limits.
public class YamlJsonSchemaTests {
    private const string DeploymentSchema = """
        {
          "type": "object",
          "required": ["name", "spec"],
          "properties": {
            "name": { "type": "string", "pattern": "^[a-z][a-z0-9-]*$" },
            "enabled": { "type": "boolean" },
            "spec": { "$ref": "#/$defs/spec" }
          },
          "additionalProperties": false,
          "$defs": {
            "spec": {
              "type": "object",
              "required": ["replicas"],
              "properties": {
                "replicas": { "type": "integer", "minimum": 1 },
                "ports": { "type": "array", "items": { "type": "integer", "maximum": 65535 } }
              }
            }
          }
        }
        """;

    [Fact]
    public void FindingsPointAtTheFailingYamlValue() {
        JsonSchema schema = JsonSchema.Parse(DeploymentSchema);
        YamlDocument document = YamlDocument.Parse("""
            name: Web_App
            spec:
              replicas: "3"
              ports: [80, 70000]
            extra: true
            """);

        JsonSchemaValidationResult result = document.Validate(schema);

        Assert.False(result.IsValid);
        Assert.Collection(result.Findings,
            name => {
                Assert.Equal("/name", name.InstancePath);
                Assert.Equal("/properties/name/pattern", name.SchemaPath);
                Assert.Equal((1, 7), (name.Line, name.Column));
                Assert.Same(document.Select("name"), name.Source);
            },
            replicas => {
                Assert.Equal("/spec/replicas", replicas.InstancePath);
                Assert.Equal("/properties/spec/$ref/properties/replicas/type", replicas.SchemaPath);
                Assert.Equal("Expected integer but found string.", replicas.Message);
                Assert.Equal((3, 13), (replicas.Line, replicas.Column));
            },
            port => {
                Assert.Equal("/spec/ports/1", port.InstancePath);
                Assert.Equal("maximum", port.Keyword);
                Assert.Equal((4, 15), (port.Line, port.Column));
            },
            extra => {
                Assert.Equal("/extra", extra.InstancePath);
                Assert.Equal("additionalProperties", extra.Keyword);
                Assert.Equal("5:8 /extra: Property 'extra' is not allowed.", extra.ToString());
            });
    }

    [Fact]
    public void YamlScalarsTakeTheTypeTheirSchemaResolves() {
        JsonSchema schema = JsonSchema.Parse("""{ "properties": { "enabled": { "type": "boolean" }, "mode": { "type": "integer" } } }""");

        // Under YAML 1.1 'yes' is a boolean and 0o-less '010' is octal 8; under the core schema 'yes' is a string.
        Assert.True(YamlDocument.Parse("%YAML 1.1\n---\nenabled: yes\nmode: 010\n").Validate(schema).IsValid);
        JsonSchemaValidationResult core = YamlDocument.Parse("enabled: yes\nmode: 0x1F\n").Validate(schema);
        JsonSchemaFinding finding = Assert.Single(core.Findings);
        Assert.Equal("/enabled", finding.InstancePath);
        Assert.Equal("Expected boolean but found string.", finding.Message);
    }

    [Fact]
    public void MergeKeysAndAliasesProvideValues() {
        JsonSchema schema = JsonSchema.Parse("""
            { "additionalProperties": { "type": "object", "required": ["image", "restart"] } }
            """);
        YamlDocument document = YamlDocument.Parse("""
            defaults: &defaults
              image: nginx
              restart: always
            web:
              <<: *defaults
              restart: unless-stopped
            worker: *defaults
            """);

        Assert.True(document.Validate(schema).IsValid);
        JsonData web = document.ToJsonData().Properties.Single(p => p.Key == "web").Value;
        Assert.True(web.TryGetProperty("image", out JsonData image));
        Assert.Equal("nginx", image.GetString());
        Assert.True(web.TryGetProperty("restart", out JsonData restart));
        Assert.Equal("unless-stopped", restart.GetString());
    }

    [Fact]
    public void EachDocumentOfAStreamValidatesSeparately() {
        JsonSchema schema = JsonSchema.Parse("""{ "required": ["kind"] }""");
        YamlDocument stream = YamlDocument.Parse("kind: Service\n---\nmetadata: {}\n---\nkind: Pod\n");

        bool[] valid = stream.Documents.Select(part => part.Validate(schema).IsValid).ToArray();

        Assert.Equal(new[] { true, false, true }, valid);
        Assert.Equal(3, stream.Documents[1].Validate(schema).Findings[0].Line);
    }

    [Fact]
    public void ReferenceCyclesStopAtTheDepthLimit() {
        JsonSchema schema = JsonSchema.Parse("""{ "$defs": { "a": { "$ref": "#/$defs/b" }, "b": { "$ref": "#/$defs/a" } }, "$ref": "#/$defs/a" }""",
            new JsonSchemaOptions { MaxDepth = 32 });

        JsonSchemaValidationResult result = schema.Validate(JsonData.Parse("{}"));

        JsonSchemaFinding finding = Assert.Single(result.Findings);
        Assert.Contains("MaxDepth", finding.Message, StringComparison.Ordinal);
        Assert.False(schema.IsValid(JsonData.Parse("{}")));
    }

    [Fact]
    public void ExponentialSchemasStopAtTheEvaluationBudget() {
        // Forty levels, each trying both branches of an anyOf that ends in false: 2^40 evaluations without the budget.
        var definitions = Enumerable.Range(0, 40).Select(i => $"\"n{i}\": {{ \"anyOf\": [ {{ \"$ref\": \"#/$defs/n{i + 1}\" }}, {{ \"$ref\": \"#/$defs/n{i + 1}\" }} ] }}");
        string schemaText = "{ \"$defs\": { " + string.Join(", ", definitions) + ", \"n40\": false }, \"$ref\": \"#/$defs/n0\" }";
        JsonSchema schema = JsonSchema.Parse(schemaText, new JsonSchemaOptions { MaxEvaluationSteps = 10_000 });

        JsonSchemaValidationResult result = schema.Validate(JsonData.Parse("1"));

        Assert.Contains(result.Findings, f => f.Message.Contains("MaxEvaluationSteps", StringComparison.Ordinal));
    }

    [Fact]
    public void CatastrophicPatternsTimeOut() {
        JsonSchema schema = JsonSchema.Parse("""{ "pattern": "^(a+)+$" }""", new JsonSchemaOptions { RegexTimeout = TimeSpan.FromMilliseconds(50) });

        JsonSchemaValidationResult result = schema.Validate(JsonData.CreateString(new string('a', 40) + "!"));

        JsonSchemaFinding finding = Assert.Single(result.Findings);
        Assert.Equal("pattern", finding.Keyword);
        Assert.Contains("timed out", finding.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FindingsStopAtMaxErrors() {
        JsonSchema schema = JsonSchema.Parse("""{ "items": { "type": "string" } }""", new JsonSchemaOptions { MaxErrors = 3 });

        JsonSchemaValidationResult result = schema.Validate(JsonData.Parse("[1, 2, 3, 4, 5]"));

        Assert.Equal(3, result.Findings.Count);
        Assert.True(result.IsTruncated);
    }

    [Fact]
    public void ExternalReferencesResolveOnlyThroughTheResolver() {
        var requested = new List<string>();
        var options = new JsonSchemaOptions {
            BaseUri = "https://example.com/schemas/app.json",
            ReferenceResolver = uri => {
                requested.Add(uri);
                return uri == "https://example.com/schemas/port.json" ? JsonData.Parse("""{ "type": "integer", "maximum": 65535 }""") : null;
            }
        };
        JsonSchema schema = JsonSchema.Parse("""{ "properties": { "port": { "$ref": "port.json" }, "host": { "$ref": "host.json" } } }""", options);

        JsonSchemaValidationResult first = schema.Validate(JsonData.Parse("""{ "port": 70000, "host": "a" }"""));
        schema.Validate(JsonData.Parse("""{ "port": 1, "host": "b" }"""));

        Assert.Equal(new[] { "https://example.com/schemas/port.json", "https://example.com/schemas/host.json" }, requested);
        Assert.Equal(new[] { "maximum", "$ref" }, first.Findings.Select(f => f.Keyword));
        Assert.Contains("https://example.com/schemas/host.json", first.Findings[1].Message, StringComparison.Ordinal);
        Assert.Throws<JsonSchemaException>(() => JsonSchema.Parse("""{ "minLength": "two" }"""));
    }

    [Fact]
    public void UnknownFormatsAreReportedAsUnchecked() {
        JsonSchema schema = JsonSchema.Parse("""{ "properties": { "when": { "format": "date-time" }, "colour": { "format": "color" } } }""");

        JsonSchemaValidationResult result = schema.Validate(JsonData.Parse("""{ "when": "2024-02-30T10:00:00Z", "colour": "red" }"""));

        Assert.Equal("format", Assert.Single(result.Findings).Keyword);
        Assert.Equal(new[] { "color" }, result.UncheckedFormats);
        JsonSchema annotationOnly = JsonSchema.Parse("""{ "format": "date" }""", new JsonSchemaOptions { ValidateFormats = false });
        Assert.True(annotationOnly.Validate(JsonData.CreateString("not a date")).IsValid);
    }

    [Fact]
    public void PatternsUseEcmaScriptCharacterClasses() {
        JsonSchema digits = JsonSchema.Parse("""{ "pattern": "^\\d+$" }""");
        JsonSchema end = JsonSchema.Parse("""{ "pattern": "^abc$" }""");

        Assert.False(digits.IsValid(JsonData.CreateString("٣٤")));
        Assert.True(digits.IsValid(JsonData.CreateString("34")));
        Assert.False(end.IsValid(JsonData.CreateString("abc\n")));
        Assert.Throws<JsonSchemaException>(() => JsonSchema.Parse("""{ "pattern": "\\a" }"""));
    }
}

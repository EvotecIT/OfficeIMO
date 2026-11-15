using System.Text;
using OfficeIMO.Json;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// Validator contracts the official suite does not reach: exact numbers at extreme exponents, shared subtrees,
// ECMA-262 regex semantics .NET differs on, relative base URIs, and bounded work for hostile schemas and values.
public class JsonSchemaEdgeCaseTests {
    [Fact]
    public void ItemsAndPropertiesCannotBeChangedThroughACast() {
        JsonData data = JsonData.Parse("{\"a\":[1]}");

        Assert.False(data.Properties is KeyValuePair<string, JsonData>[]);
        Assert.False(data.Properties[0].Value.Items is JsonData[]);
    }

    [Fact]
    public void NumbersStayExactAtExtremeExponents() {
        Assert.False(JsonData.DeepEquals(JsonData.Parse("1e400"), JsonData.CreateNumber(double.PositiveInfinity)));
        Assert.Throws<FormatException>(() => JsonData.Parse("1e1000000000"));

        Assert.True(JsonSchema.Parse("{\"multipleOf\":1}").IsValid(JsonData.Parse("1e10001")));
        Assert.False(JsonSchema.Parse("{\"multipleOf\":1}").IsValid(JsonData.Parse("1e-10001")));
        JsonSchema tiny = JsonSchema.Parse("{\"multipleOf\":1e-400}");
        Assert.True(tiny.IsValid(JsonData.Parse("2e-400")));
        Assert.False(tiny.IsValid(JsonData.Parse("1.5e-400")));
    }

    [Fact]
    public void SharedSubtreesCompareHashAndPreviewInLinearTime() {
        JsonData x = JsonData.CreateNumber(1);
        JsonData y = JsonData.CreateNumber(1);
        for (int i = 0; i < 40; i++) {
            x = JsonData.CreateArray(new[] { x, x });
            y = JsonData.CreateArray(new[] { y, y });
        }

        Assert.True(JsonData.DeepEquals(x, y));
        Assert.False(JsonSchema.Parse("{\"uniqueItems\":true}").IsValid(JsonData.CreateArray(new[] { x, JsonData.CreateNumber(0), y })));
        JsonSchema schema = JsonSchema.Load(JsonData.CreateObject(new[] { new KeyValuePair<string, JsonData>("const", x) }));
        JsonSchemaFinding finding = Assert.Single(schema.Validate(JsonData.CreateNumber(0)).Findings);
        Assert.True(finding.Message.Length < 200, finding.Message);
    }

    [Fact]
    public void PreviewsOfLongStringsAndNamesStayShort() {
        string huge = new string('x', 1_000_000);
        JsonData[] constants = {
            JsonData.CreateString(huge),
            JsonData.CreateObject(new[] { new KeyValuePair<string, JsonData>(huge, JsonData.CreateNull()) })
        };

        foreach (JsonData constant in constants) {
            JsonSchema schema = JsonSchema.Load(JsonData.CreateObject(new[] { new KeyValuePair<string, JsonData>("const", constant) }));
            JsonSchemaFinding finding = Assert.Single(schema.Validate(JsonData.CreateNumber(0)).Findings);
            Assert.True(finding.Message.Length < 200, finding.Message.Length.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
    }

    [Theory]
    [InlineData("^..$", "\U0001F600", false)]
    [InlineData("^.$", "\U0001F600", true)]
    [InlineData("(?s:^.$)", "\n", true)]
    [InlineData("^.$", "\n", false)]
    [InlineData("(?m:^a$)", "a\nb", true)]
    [InlineData("^a$", "a\n", false)]
    [InlineData("^\\w+\\b", "aé", true)]
    [InlineData("^(?<a>a)(b)\\1$", "aba", true)]
    [InlineData("^(?<a>a)(b)\\1$", "abb", false)]
    [InlineData("^(?<a>a)(b)\\k<a>$", "aba", true)]
    [InlineData("^(?:(?<x>a)|(?<x>b))(c)\\2$", "bcb", true)]
    [InlineData("^(?:(?<x>a)|(?<x>b))\\k<x>$", "bb", true)]
    [InlineData("^(?:(?<x>a)|(?<x>b))\\k<x>$", "ba", false)]
    [InlineData("^(a)?\\1b$", "b", true)]
    public void PatternsFollowEcmaScriptSemantics(string pattern, string value, bool valid) {
        JsonSchema schema = JsonSchema.Load(JsonData.CreateObject(new[] { new KeyValuePair<string, JsonData>("pattern", JsonData.CreateString(pattern)) }));

        Assert.Equal(valid, schema.IsValid(JsonData.CreateString(value)));
    }

    [Fact]
    public void UnpairedSurrogatesFailIdnHostname() {
        JsonSchema schema = JsonSchema.Parse("{\"format\":\"idn-hostname\"}");

        Assert.False(schema.IsValid(JsonData.CreateString("a\uD800b")));
    }

    [Fact]
    public void RelativeReferencesResolveAgainstARelativeBase() {
        var requested = new List<string>();
        var options = new JsonSchemaOptions {
            BaseUri = "schemas/root.json",
            ReferenceResolver = uri => {
                requested.Add(uri);
                return JsonData.Parse("{\"type\":\"string\"}");
            }
        };

        Assert.False(JsonSchema.Parse("{\"$ref\":\"child.json\"}", options).IsValid(JsonData.CreateNumber(1)));
        Assert.Equal(new[] { "schemas/child.json" }, requested);
    }

    [Fact]
    public void ValuesCannotBeBuiltDeeperThanTheNestingLimit() {
        JsonData value = JsonData.CreateNumber(1);
        for (int i = 0; i < 1024; i++) value = JsonData.CreateArray(new[] { value });

        Assert.Throws<ArgumentException>(() => JsonData.CreateArray(new[] { value }));
        Assert.Throws<ArgumentException>(() => JsonData.CreateObject(new[] { new KeyValuePair<string, JsonData>("a", value) }));
    }

    [Fact]
    public void UnresolvedReferencesFailInsideNot() {
        JsonSchema schema = JsonSchema.Parse("{\"not\":{\"$ref\":\"missing.json\"}}");

        JsonSchemaValidationResult result = schema.Validate(JsonData.CreateNumber(1));

        Assert.False(result.IsValid);
        Assert.Equal("$ref", Assert.Single(result.Findings).Keyword);
    }

    [Fact]
    public void ASchemaValueSharedByTwoResourcesResolvesAgainstEach() {
        JsonData shared = JsonData.Parse("{\"$ref\":\"child.json\"}");
        JsonData Resource(string id) => JsonData.CreateObject(new[] {
            new KeyValuePair<string, JsonData>("$id", JsonData.CreateString(id)),
            new KeyValuePair<string, JsonData>("properties", JsonData.CreateObject(new[] { new KeyValuePair<string, JsonData>("v", shared) }))
        });
        JsonData root = JsonData.CreateObject(new[] {
            new KeyValuePair<string, JsonData>("$defs", JsonData.CreateObject(new[] {
                new KeyValuePair<string, JsonData>("a", Resource("https://example.test/a/")),
                new KeyValuePair<string, JsonData>("b", Resource("https://example.test/b/"))
            })),
            new KeyValuePair<string, JsonData>("$ref", JsonData.CreateString("https://example.test/b/"))
        });
        var options = new JsonSchemaOptions {
            ReferenceResolver = uri => JsonData.Parse(uri == "https://example.test/a/child.json" ? "{\"type\":\"string\"}" : "{\"type\":\"number\"}")
        };

        Assert.True(JsonSchema.Load(root, options).IsValid(JsonData.Parse("{\"v\":1}")));
    }

    [Theory]
    [InlineData("(?<>a)")]
    [InlineData("(?<1a>a)")]
    [InlineData("(?<a>a)\\k<>")]
    public void PatternsWithInvalidGroupNamesFailToLoad(string pattern) {
        JsonData schema = JsonData.CreateObject(new[] { new KeyValuePair<string, JsonData>("pattern", JsonData.CreateString(pattern)) });

        Assert.Throws<JsonSchemaException>(() => JsonSchema.Load(schema));
    }

    [Fact]
    public void AMalformedBaseUriFailsToLoad() {
        JsonSchemaException error = Assert.Throws<JsonSchemaException>(() =>
            JsonSchema.Parse("{\"$ref\":\"child.json\"}", new JsonSchemaOptions { BaseUri = "http://[invalid" }));

        Assert.Contains("BaseUri", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("schemas/root.json")]
    [InlineData(null)]
    public void DotSegmentsInRelativeReferencesFindEmbeddedResources(string? baseUri) {
        JsonSchema schema = JsonSchema.Parse("{\"$ref\":\"./sub/../child.json\",\"$defs\":{\"c\":{\"$id\":\"child.json\",\"type\":\"string\"}}}",
            new JsonSchemaOptions { BaseUri = baseUri });

        JsonSchemaValidationResult result = schema.Validate(JsonData.CreateNumber(1));

        Assert.Equal("type", Assert.Single(result.Findings).Keyword);
    }

    [Fact]
    public void YamlIntegersBeyondDecimalRangeStayExact() {
        JsonSchema schema = JsonSchema.Parse("{\"properties\":{\"n\":{\"const\":79228162514264337593543950337}}}");

        Assert.True(YamlDocument.Parse("n: 79228162514264337593543950337\n").Validate(schema).IsValid);
        Assert.False(YamlDocument.Parse("n: 79228162514264337593543950338\n").Validate(schema).IsValid);
    }

    [Fact]
    public void AReferencedSchemaThatFailsToCompileFailsEveryValidation() {
        var options = new JsonSchemaOptions { ReferenceResolver = _ => JsonData.Parse("{\"type\":\"string\",\"minLength\":-1}") };
        JsonSchema schema = JsonSchema.Parse("{\"$ref\":\"https://example.com/bad.json\"}", options);

        Assert.Throws<JsonSchemaException>(() => schema.Validate(JsonData.CreateString("abc")));
        Assert.Throws<JsonSchemaException>(() => schema.Validate(JsonData.CreateString("abc")));
    }

    [Fact]
    public void ReusedAnchorsCountTheirFullDepth() {
        var yaml = new StringBuilder("n0: &n0 [0]\n");
        for (int i = 1; i <= 600; i++) yaml.Append('n').Append(i).Append(": &n").Append(i).Append(" [*n").Append(i - 1).Append("]\n");

        // Loading allows the chain; converting it must still stop at 512 levels.
        YamlDocument document = YamlDocument.Parse(yaml.ToString(), new YamlLoadOptions { MaxDepth = 10_000, MaxAliasExpansion = 10_000_000 });

        Assert.Throws<InvalidOperationException>(() => document.ToJsonData());
    }
}

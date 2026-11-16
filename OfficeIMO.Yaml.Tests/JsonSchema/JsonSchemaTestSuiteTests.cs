using OfficeIMO.Json;
using OfficeIMO.Tests;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Yaml.Tests;

// The validator's main proof: the official JSON-Schema-Test-Suite (OfficeIMO.TestAssets/Json/JSON-Schema-Test-Suite),
// run for draft 7 and draft 2020-12 with the suite's compatibility rules. Known gaps are listed with a reason; a listed
// test that starts passing fails the run, so the list only shrinks.
public class JsonSchemaTestSuiteTests {
    private readonly ITestOutputHelper _output;

    public JsonSchemaTestSuiteTests(ITestOutputHelper output) => _output = output;

    private static string Directory => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Json", "JSON-Schema-Test-Suite", "tests");

    /// <summary>Known failures, keyed as <c>dialect: file | case | test</c>, with the reason.</summary>
    private static readonly IReadOnlyDictionary<string, string> Exceptions = new Dictionary<string, string>(StringComparer.Ordinal) {
        ["2020: optional/cross-draft.json | refs to historic drafts are processed as historic drafts (2019) | first item not a string is valid"] = "The dialect is chosen once for the whole schema; a $schema in a referenced resource does not switch it.",
        ["2020: optional/cross-draft.json | refs to historic drafts are processed as historic drafts (7) | missing bar is valid"] = "The dialect is chosen once for the whole schema; a $schema in a referenced resource does not switch it.",
        ["2020: vocabulary.json | schema that uses custom metaschema with no validation vocabulary | no validation: invalid number, but it still validates"] = "Custom meta-schemas and $vocabulary are not read; every keyword is always active.",
        ["7: optional/content.json | validation of binary string-encoding | an invalid base64 string (% is not a valid character)"] = "contentEncoding and contentMediaType are annotations only, as draft 2019-09 and later require.",
        ["7: optional/content.json | validation of binary-encoded media type documents | a validly-encoded invalid JSON document"] = "contentEncoding and contentMediaType are annotations only, as draft 2019-09 and later require.",
        ["7: optional/content.json | validation of binary-encoded media type documents | an invalid base64 string that is valid JSON"] = "contentEncoding and contentMediaType are annotations only, as draft 2019-09 and later require.",
        ["7: optional/content.json | validation of string-encoded content based on media type | an invalid JSON document"] = "contentEncoding and contentMediaType are annotations only, as draft 2019-09 and later require.",
        ["7: optional/cross-draft.json | refs to future drafts are processed as future drafts (2019) | missing bar is invalid"] = "The dialect is chosen once for the whole schema; a $schema in a referenced resource does not switch it.",
        ["7: optional/cross-draft.json | refs to future drafts are processed as future drafts (2020) | first item not a string is invalid"] = "The dialect is chosen once for the whole schema; a $schema in a referenced resource does not switch it."
    };

    [Theory]
    [InlineData(7)]
    [InlineData(2020)]
    public void PassesTheOfficialSuite(int dialect) {
        var failures = new List<string>();
        var stale = new List<string>();
        int required = 0, requiredPassed = 0, optional = 0, optionalPassed = 0;
        foreach (string path in System.IO.Directory.GetFiles(Directory, "*.json", SearchOption.AllDirectories).OrderBy(p => p, StringComparer.Ordinal)) {
            string file = path.Substring(Directory.Length + 1).Replace(Path.DirectorySeparatorChar, '/');
            bool isOptional = file.StartsWith("optional/", StringComparison.Ordinal);
            JsonData suite = JsonData.Parse(File.ReadAllText(path));
            suite.TryGetProperty("tests", out JsonData cases);
            foreach (JsonData testCase in cases.Items) {
                if (!IsCompatible(testCase, dialect)) continue;
                string[] tags = testCase.TryGetProperty("tags", out JsonData tagList) ? tagList.Items.Select(t => t.GetString()).ToArray() : Array.Empty<string>();
                string caseName = testCase.TryGetProperty("description", out JsonData caseDescription) ? caseDescription.GetString() : "?";
                var options = new JsonSchemaOptions {
                    Draft = dialect == 7 ? JsonSchemaDraft.Draft7 : JsonSchemaDraft.Draft202012,
                    ValidateFormats = !tags.Contains("format-annotation")
                };
                if (testCase.TryGetProperty("externalSchemas", out JsonData external)) {
                    options.ReferenceResolver = uri => external.TryGetProperty(uri, out JsonData schema) ? schema : null;
                }
                testCase.TryGetProperty("schema", out JsonData schemaData);
                JsonSchema? schema = null;
                string? loadError = null;
                try {
                    schema = JsonSchema.Load(schemaData, options);
                } catch (JsonSchemaException error) {
                    loadError = error.Message;
                }
                testCase.TryGetProperty("tests", out JsonData tests);
                foreach (JsonData test in tests.Items) {
                    test.TryGetProperty("description", out JsonData testDescription);
                    test.TryGetProperty("data", out JsonData data);
                    test.TryGetProperty("valid", out JsonData expected);
                    string id = file + " | " + caseName + " | " + testDescription.GetString();
                    string? actual;
                    try {
                        actual = schema is null ? "load failed: " + loadError
                            : schema.Validate(data).IsValid == expected.GetBoolean() ? null
                            : "expected " + (expected.GetBoolean() ? "valid" : "invalid");
                    } catch (Exception error) when (error is JsonSchemaException or InvalidOperationException) {
                        actual = error.GetType().Name + ": " + error.Message;
                    }
                    if (isOptional) optional++; else required++;
                    string key = dialect.ToString(System.Globalization.CultureInfo.InvariantCulture) + ": " + id;
                    if (actual is null) {
                        if (isOptional) optionalPassed++; else requiredPassed++;
                        if (Exceptions.ContainsKey(key)) stale.Add(key);
                    } else if (!Exceptions.ContainsKey(key)) {
                        failures.Add(key + " -> " + actual);
                    }
                }
            }
        }
        _output.WriteLine($"Draft {dialect}: required {requiredPassed}/{required}, optional {optionalPassed}/{optional}.");
        foreach (string failure in failures) _output.WriteLine(failure);
        Assert.True(failures.Count == 0, failures.Count + " unexpected failures:\n" + string.Join("\n", failures));
        Assert.True(stale.Count == 0, "Listed failures now pass; remove them:\n" + string.Join("\n", stale));
    }

    /// <summary>The suite's compatibility rule: constraints form a disjunction over releases, evaluated in order.</summary>
    private static bool IsCompatible(JsonData testCase, int dialect) {
        if (!testCase.TryGetProperty("compatibility", out JsonData compatibility)) return true;
        bool valid = true;
        foreach (string constraint in compatibility.GetString().Split(',')) {
            string op = constraint.StartsWith("<=", StringComparison.Ordinal) ? "<=" : constraint.StartsWith("=", StringComparison.Ordinal) ? "=" : string.Empty;
            int version = int.Parse(constraint.Substring(op.Length), System.Globalization.CultureInfo.InvariantCulture);
            valid = op switch {
                "<=" => dialect <= version,
                "=" => dialect == version,
                _ => dialect >= version
            };
            if (dialect <= version) break;
        }
        return valid;
    }
}

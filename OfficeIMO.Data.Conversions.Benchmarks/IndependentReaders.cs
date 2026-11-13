using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using CsvHelper;
using Newtonsoft.Json.Linq;
using OfficeIMO.Tests;
using OfficeIMO.Yaml.Tests.Conformance;

namespace OfficeIMO.Data.Conversions.Benchmarks;

/// <summary>
/// Opt-in check (<c>--verify</c>) that established libraries read OfficeIMO's output to the same data: YamlDotNet reads
/// the YAML, Newtonsoft.Json reads the JSON and the XML (through its own XML convention), CsvHelper reads the CSV and
/// Tomlyn reads the TOML.
/// Inputs are the pinned JSONTestSuite accept cases, the yaml-test-suite in.json values and generated records with
/// awkward text. Each disagreement is printed; the exit code is the number of failures.
/// </summary>
internal static class IndependentReaders {
    public static int Run() {
        // YamlDotNet parses numbers with the current culture; a comma-decimal culture would turn 0.5 into text.
        CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
        var failures = new List<string>();
        List<(string Name, byte[] Json)> inputs = JsonInputs();
        int yaml = 0, json = 0, xml = 0, csv = 0, toml = 0, tomlReported = 0;
        foreach ((string name, byte[] input) in inputs) {
            using JsonDocument expected = JsonDocument.Parse(input);
            string yamlText = Convert(input, DataFormat.Json, DataFormat.Yaml);
            Check(failures, name + " (YamlDotNet)", Same(expected.RootElement, YamlDotNetValue(yamlText)), yamlText);
            yaml++;
            JToken newtonsoft = Parse(Convert(input, DataFormat.Json, DataFormat.Json));
            Check(failures, name + " (Newtonsoft JSON)", Same(expected.RootElement, ToPlain(newtonsoft)), null);
            json++;

            // TOML: Tomlyn reads the data back, less what the report names (nulls are omitted, a non-table root is
            // wrapped in "root"). Numbers TOML cannot hold exactly are reported and left out of the comparison.
            (string tomlText, OfficeConversionReport report) = ConvertWithReport(input, DataFormat.Json, DataFormat.Toml);
            if (report.Diagnostics.Any(d => d.Code is OfficeDiagnosticCodes.NumberOutOfRange or OfficeDiagnosticCodes.NumberPrecision)) {
                tomlReported++;
                continue;
            }

            object? tomlyn;
            try {
                tomlyn = TomlynPlain(Tomlyn.TomlSerializer.Deserialize<Tomlyn.Model.TomlTable>(tomlText));
            } catch (Exception error) {
                failures.Add(name + " (Tomlyn): " + error.Message + "\n" + tomlText);
                continue;
            }

            using JsonDocument withoutNulls = JsonDocument.Parse(TomlShape(expected.RootElement));
            Check(failures, name + " (Tomlyn)", Same(withoutNulls.RootElement, tomlyn), tomlText);
            toml++;
        }

        // XML: Newtonsoft's SerializeXNode and OfficeIMO's reader share the convention for data XML, so both must agree.
        foreach (string text in new[] { Workloads.NestedJson(200), Workloads.FlatJson(200), Awkward() }) {
            byte[] input = Encoding.UTF8.GetBytes(text);
            string ours = Convert(input, DataFormat.Json, DataFormat.Xml);
            JToken newtonsoft = Parse(Newtonsoft.Json.JsonConvert.SerializeXNode(XDocument.Parse(ours)));
            // The two documented convention differences: Newtonsoft keeps the declaration as "?xml" and reads an empty
            // element as null, where OfficeIMO reads "" (it writes null as xsi:nil, which these inputs do not contain).
            ((JObject)newtonsoft).Remove("?xml");
            foreach (JValue empty in newtonsoft.SelectTokens("..*").OfType<JValue>().Where(v => v.Type == JTokenType.Null).ToList()) empty.Value = string.Empty;
            using JsonDocument back = JsonDocument.Parse(Convert(Encoding.UTF8.GetBytes(ours), DataFormat.Xml, DataFormat.Json));
            bool same = Same(back.RootElement, ToPlain(newtonsoft));
            if (!same && Environment.GetEnvironmentVariable("OFFICEIMO_VERIFY_DUMP") is { } dump) {
                File.WriteAllText(Path.Combine(dump, "ours-" + xml + ".json"), back.RootElement.GetRawText());
                File.WriteAllText(Path.Combine(dump, "newtonsoft-" + xml + ".json"), newtonsoft.ToString(Newtonsoft.Json.Formatting.None));
            }

            Check(failures, "XML " + text.Length.ToString(CultureInfo.InvariantCulture) + " chars (Newtonsoft XML)", same, null);
            xml++;
        }

        // CSV: every field CsvHelper reads equals the source value's text.
        foreach (string text in new[] { Workloads.FlatJson(500), Awkward() }) {
            byte[] input = Encoding.UTF8.GetBytes(text);
            string ours = Convert(input, DataFormat.Json, DataFormat.Csv);
            using JsonDocument source = JsonDocument.Parse(input);
            using var reader = new CsvReader(new StringReader(ours), CultureInfo.InvariantCulture);
            reader.Read();
            reader.ReadHeader();
            int row = 0;
            foreach (JsonElement record in source.RootElement.EnumerateArray()) {
                reader.Read();
                foreach (JsonProperty property in record.EnumerateObject()) {
                    string expectedText = property.Value.ValueKind == JsonValueKind.String ? property.Value.GetString()! : property.Value.GetRawText();
                    string? actual = reader.GetField(property.Name);
                    if (actual != expectedText) failures.Add("CSV row " + row + " field " + property.Name + ": expected '" + expectedText + "' but CsvHelper read '" + actual + "'");
                }

                row++;
            }

            csv++;
        }

        Console.WriteLine($"YamlDotNet: {yaml} documents, Newtonsoft JSON: {json}, Newtonsoft XML: {xml}, CsvHelper: {csv} files, Tomlyn: {toml} documents ({tomlReported} with reported loss skipped); {failures.Count} disagreements.");
        foreach (string failure in failures.Take(50)) Console.WriteLine("  " + failure);
        return failures.Count;
    }

    private static List<(string, byte[])> JsonInputs() {
        var inputs = new List<(string, byte[])>();
        string suite = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Json", "JSONTestSuite", "test_parsing");
        foreach (string path in Directory.GetFiles(suite, "y_*.json").OrderBy(p => p, StringComparer.Ordinal)) {
            string name = Path.GetFileName(path);
            if (name.Contains("duplicated_key", StringComparison.Ordinal)) continue;
            inputs.Add((name, File.ReadAllBytes(path)));
        }

        foreach (YamlTestSuiteCase testCase in YamlTestSuite.Load().Cases) {
            if (testCase.IsError || testCase.Json == null) continue;
            IReadOnlyList<YamlConformanceNode> documents = YamlJsonReader.ReadStream(testCase.Json);
            if (documents.Count != 1) continue;
            inputs.Add(("yaml-test-suite " + testCase.Id, Encoding.UTF8.GetBytes(Convert(testCase.Input, DataFormat.Yaml, DataFormat.Json))));
        }

        inputs.Add(("awkward records", Encoding.UTF8.GetBytes(Awkward())));
        return inputs;
    }

    // Text that trips YAML 1.1 readers and CSV writers.
    private static string Awkward() =>
        "[{\"a\":\"yes\",\"b\":\"012\",\"c\":\"1:20\",\"d\":\"line\\nbreak\",\"e\":\"comma, \\\"quote\\\"\",\"f\":\"~\",\"g\":\"null\",\"h\":\" padded \",\"i\":\"#hash\",\"j\":\"- dash\"}," +
        "{\"a\":\"on\",\"b\":\"0x1F\",\"c\":\".inf\",\"d\":\"tab\\there\",\"e\":\"=SUM(A1)\",\"f\":\"\",\"g\":\"True\",\"h\":\"2026-11-10\",\"i\":\"@at\",\"j\":\"key: value\"}]";

    private static string Convert(byte[] input, DataFormat source, DataFormat target) {
        using var output = new MemoryStream();
        DataConversionEngine.Convert(new MemoryStream(input), source, output, target, null);
        return Encoding.UTF8.GetString(output.ToArray());
    }

    private static (string, OfficeConversionReport) ConvertWithReport(byte[] input, DataFormat source, DataFormat target) {
        using var output = new MemoryStream();
        OfficeConversionReport report = DataConversionEngine.Convert(new MemoryStream(input), source, output, target, null);
        return (Encoding.UTF8.GetString(output.ToArray()), report);
    }

    // The data TOML holds: nulls removed and a root that is not an object wrapped in "root".
    private static string TomlShape(JsonElement root) {
        using var output = new MemoryStream();
        using (var writer = new Utf8JsonWriter(output)) {
            if (root.ValueKind == JsonValueKind.Null) {
                writer.WriteStartObject();
                writer.WriteEndObject();
            } else if (root.ValueKind != JsonValueKind.Object) {
                writer.WriteStartObject();
                writer.WritePropertyName("root");
                WithoutNulls(writer, root);
                writer.WriteEndObject();
            } else {
                WithoutNulls(writer, root);
            }
        }

        return Encoding.UTF8.GetString(output.ToArray());
    }

    private static void WithoutNulls(Utf8JsonWriter writer, JsonElement element) {
        switch (element.ValueKind) {
            case JsonValueKind.Object:
                writer.WriteStartObject();
                foreach (JsonProperty property in element.EnumerateObject()) {
                    if (property.Value.ValueKind == JsonValueKind.Null) continue;
                    writer.WritePropertyName(property.Name);
                    WithoutNulls(writer, property.Value);
                }

                writer.WriteEndObject();
                break;
            case JsonValueKind.Array:
                writer.WriteStartArray();
                foreach (JsonElement item in element.EnumerateArray()) {
                    if (item.ValueKind != JsonValueKind.Null) WithoutNulls(writer, item);
                }

                writer.WriteEndArray();
                break;
            case JsonValueKind.Null:
                break;
            default:
                element.WriteTo(writer);
                break;
        }
    }

    private static object? TomlynPlain(object? value) => value switch {
        Tomlyn.Model.TomlTable table => table.ToDictionary(entry => (object)entry.Key, entry => TomlynPlain(entry.Value)),
        Tomlyn.Model.TomlTableArray tables => tables.Select(item => TomlynPlain(item)).ToList(),
        Tomlyn.Model.TomlArray array => array.Select(TomlynPlain).ToList(),
        _ => value
    };

    private static object? YamlDotNetValue(string yaml) {
        var deserializer = new YamlDotNet.Serialization.DeserializerBuilder().WithAttemptingUnquotedStringTypeDeserialization().Build();
        return deserializer.Deserialize(new StringReader(yaml));
    }

    // Dates stay text, as they are in the JSON; Newtonsoft would otherwise turn them into DateTime values.
    private static JToken Parse(string json) {
        using var reader = new Newtonsoft.Json.JsonTextReader(new StringReader(json)) { DateParseHandling = Newtonsoft.Json.DateParseHandling.None };
        return JToken.ReadFrom(reader);
    }

    private static object? ToPlain(JToken token) => token.Type switch {
        JTokenType.Object => ((JObject)token).Properties().ToDictionary(p => (object)p.Name, p => ToPlain(p.Value)),
        JTokenType.Array => ((JArray)token).Select(ToPlain).ToList(),
        JTokenType.Null => null,
        JTokenType.Boolean => token.Value<bool>(),
        JTokenType.Integer or JTokenType.Float => double.Parse(token.ToString(Newtonsoft.Json.Formatting.None), CultureInfo.InvariantCulture),
        _ => token.Value<string>()
    };

    // Compares a JSON value with a plain object graph: dictionaries, lists, strings, booleans and numbers by value.
    private static bool Same(JsonElement expected, object? actual) {
        switch (expected.ValueKind) {
            case JsonValueKind.Object:
                if (actual is not System.Collections.IDictionary map) return false;
                int count = 0;
                foreach (JsonProperty property in expected.EnumerateObject()) {
                    count++;
                    object? key = map.Keys.Cast<object>().FirstOrDefault(k => System.Convert.ToString(k, CultureInfo.InvariantCulture) == property.Name);
                    if (key == null || !Same(property.Value, map[key])) return false;
                }

                return count == map.Count;
            case JsonValueKind.Array:
                if (actual is not System.Collections.IList list) return false;
                List<JsonElement> items = expected.EnumerateArray().ToList();
                return items.Count == list.Count && items.Select((item, i) => Same(item, list[i])).All(x => x);
            case JsonValueKind.String:
                return actual is string text && text == expected.GetString();
            case JsonValueKind.Number: {
                double number = double.Parse(expected.GetRawText(), CultureInfo.InvariantCulture);
                // YamlDotNet picks the narrowest type that parses, so a float compares at single precision.
                if (actual is float single) return single == (float)number;
                return actual is not (string or bool or null) && System.Convert.ToDouble(actual, CultureInfo.InvariantCulture) == number;
            }
            case JsonValueKind.True:
            case JsonValueKind.False:
                return actual is bool flag && flag == expected.GetBoolean();
            default:
                return actual == null;
        }
    }

    private static void Check(List<string> failures, string name, bool same, string? output) {
        if (!same) failures.Add(name + (output == null ? string.Empty : ":\n" + output));
    }
}

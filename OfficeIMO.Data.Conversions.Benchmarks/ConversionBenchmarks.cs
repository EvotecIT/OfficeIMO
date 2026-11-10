using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using BenchmarkDotNet.Attributes;
using CsvHelper;
using Newtonsoft.Json;
using YamlDotNet.Core;
using YamlDotNet.Core.Events;

namespace OfficeIMO.Data.Conversions.Benchmarks;

/// <summary>
/// Each lane converts the same input with OfficeIMO and with the common pipeline built from established libraries.
/// Setup validates both outputs once, so a lane cannot win by doing less work.
/// </summary>
[MemoryDiagnoser]
public class ConversionBenchmarks {
    private byte[] _yaml = Array.Empty<byte>();
    private string _yamlText = string.Empty;
    private byte[] _xml = Array.Empty<byte>();
    private string _xmlText = string.Empty;
    private byte[] _nestedJson = Array.Empty<byte>();
    private string _nestedJsonText = string.Empty;
    private byte[] _flatJson = Array.Empty<byte>();
    private byte[] _csv = Array.Empty<byte>();
    private string _csvText = string.Empty;

    [Params(20_000)]
    public int Records { get; set; }

    [GlobalSetup]
    public void Setup() {
        _yamlText = Workloads.Yaml(Records / 4);
        _yaml = Encoding.UTF8.GetBytes(_yamlText);
        _xmlText = Workloads.Xml(Records);
        _xml = Encoding.UTF8.GetBytes(_xmlText);
        _nestedJsonText = Workloads.NestedJson(Records);
        _nestedJson = Encoding.UTF8.GetBytes(_nestedJsonText);
        _flatJson = Encoding.UTF8.GetBytes(Workloads.FlatJson(Records));
        _csvText = Workloads.Csv(Records);
        _csv = Encoding.UTF8.GetBytes(_csvText);
        Validation.Run(this);
    }

    // ---------------------------------------------------------------- YAML -> JSON

    [Benchmark(Description = "YAML->JSON OfficeIMO")]
    public long YamlToJsonOfficeIMO() => Run(_yaml, DataFormat.Yaml, DataFormat.Json);

    [Benchmark(Description = "YAML->JSON YamlDotNet (JSON-compatible serializer)")]
    public long YamlToJsonYamlDotNet() {
        var deserializer = new YamlDotNet.Serialization.DeserializerBuilder().Build();
        var serializer = new YamlDotNet.Serialization.SerializerBuilder().JsonCompatible().Build();
        var parser = new YamlDotNet.Core.Parser(new StringReader(_yamlText));
        parser.Consume<StreamStart>();
        var output = new StringWriter();
        output.Write('[');
        bool first = true;
        while (parser.Accept<DocumentStart>(out _)) {
            if (!first) output.Write(',');
            first = false;
            serializer.Serialize(output, deserializer.Deserialize(parser));
        }

        output.Write(']');
        return output.GetStringBuilder().Length;
    }

    // ---------------------------------------------------------------- XML <-> JSON

    [Benchmark(Description = "XML->JSON OfficeIMO")]
    public long XmlToJsonOfficeIMO() => Run(_xml, DataFormat.Xml, DataFormat.Json);

    [Benchmark(Description = "XML->JSON Newtonsoft SerializeXNode")]
    public long XmlToJsonNewtonsoft() => JsonConvert.SerializeXNode(XDocument.Parse(_xmlText), Formatting.Indented).Length;

    [Benchmark(Description = "JSON->XML OfficeIMO")]
    public long JsonToXmlOfficeIMO() => Run(_nestedJson, DataFormat.Json, DataFormat.Xml);

    [Benchmark(Description = "JSON->XML Newtonsoft DeserializeXNode")]
    public long JsonToXmlNewtonsoft() => JsonConvert.DeserializeXNode("{\"root\":{\"item\":" + _nestedJsonText + "}}")!.ToString().Length;

    // ---------------------------------------------------------------- CSV <-> JSON

    [Benchmark(Description = "JSON->CSV OfficeIMO")]
    public long JsonToCsvOfficeIMO() => Run(_flatJson, DataFormat.Json, DataFormat.Csv);

    [Benchmark(Description = "JSON->CSV JsonDocument + CsvHelper")]
    public long JsonToCsvCsvHelper() {
        using JsonDocument document = JsonDocument.Parse(_flatJson);
        var output = new StringWriter();
        using var csv = new CsvWriter(output, CultureInfo.InvariantCulture);
        bool header = true;
        foreach (JsonElement row in document.RootElement.EnumerateArray()) {
            if (header) {
                foreach (JsonProperty property in row.EnumerateObject()) csv.WriteField(property.Name);
                csv.NextRecord();
                header = false;
            }

            foreach (JsonProperty property in row.EnumerateObject()) {
                csv.WriteField(property.Value.ValueKind == JsonValueKind.String ? property.Value.GetString() : property.Value.GetRawText());
            }

            csv.NextRecord();
        }

        csv.Flush();
        return output.GetStringBuilder().Length;
    }

    [Benchmark(Description = "CSV->JSON OfficeIMO (typed columns)")]
    public long CsvToJsonOfficeIMO() => Run(_csv, DataFormat.Csv, DataFormat.Json);

    [Benchmark(Description = "CSV->JSON CsvHelper dynamic + System.Text.Json (strings only)")]
    public long CsvToJsonCsvHelper() {
        using var reader = new CsvReader(new StringReader(_csvText), CultureInfo.InvariantCulture);
        var rows = new List<IDictionary<string, object>>();
        foreach (dynamic record in reader.GetRecords<dynamic>()) rows.Add((IDictionary<string, object>)record);
        return System.Text.Json.JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented = true }).Length;
    }

    internal static long Run(byte[] input, DataFormat source, DataFormat target) {
        using var output = new MemoryStream();
        DataConversionEngine.Convert(new MemoryStream(input, writable: false), source, output, target, null);
        return output.Length;
    }

    internal byte[] YamlBytes => _yaml;
    internal byte[] XmlBytes => _xml;
    internal byte[] NestedJsonBytes => _nestedJson;
    internal byte[] FlatJsonBytes => _flatJson;
    internal byte[] CsvBytes => _csv;
}

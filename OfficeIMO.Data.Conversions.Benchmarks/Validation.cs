using System.Text;
using System.Text.Json;
using System.Xml.Linq;

namespace OfficeIMO.Data.Conversions.Benchmarks;

/// <summary>Checks once that OfficeIMO's outputs hold the same records the comparison pipelines produce.</summary>
internal static class Validation {
    public static void Run(ConversionBenchmarks lanes) {
        Require(Count(Json(lanes.YamlBytes, DataFormat.Yaml)) == lanes.Records / 4, "YAML->JSON document count");
        using (JsonDocument xml = JsonDocument.Parse(Output(lanes.XmlBytes, DataFormat.Xml, DataFormat.Json))) {
            Require(xml.RootElement.GetProperty("customers").GetProperty("customer").GetArrayLength() == lanes.Records, "XML->JSON customers");
        }

        XDocument fromJson = XDocument.Parse(Encoding.UTF8.GetString(Output(lanes.NestedJsonBytes, DataFormat.Json, DataFormat.Xml)));
        Require(fromJson.Root!.Elements("item").Count() == lanes.Records, "JSON->XML items");
        string csv = Encoding.UTF8.GetString(Output(lanes.FlatJsonBytes, DataFormat.Json, DataFormat.Csv));
        Require(csv.Split('\n').Length - 2 == lanes.Records, "JSON->CSV rows");
        Require(Count(Output(lanes.CsvBytes, DataFormat.Csv, DataFormat.Json)) == lanes.Records, "CSV->JSON rows");
    }

    private static byte[] Json(byte[] input, DataFormat source) => Output(input, source, DataFormat.Json);

    private static byte[] Output(byte[] input, DataFormat source, DataFormat target) {
        using var output = new MemoryStream();
        DataConversionEngine.Convert(new MemoryStream(input), source, output, target, null);
        return output.ToArray();
    }

    private static int Count(byte[] json) {
        using JsonDocument document = JsonDocument.Parse(json);
        return document.RootElement.GetArrayLength();
    }

    private static void Require(bool condition, string what) {
        if (!condition) throw new InvalidOperationException("Benchmark validation failed: " + what);
    }
}

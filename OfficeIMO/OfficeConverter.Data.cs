using OfficeIMO.Data.Conversions;

namespace OfficeIMO;

public static partial class OfficeConverter {
    // Structured-data routes from the capability catalog, executed by OfficeIMO.Data.Conversions and its Excel adapter.
    private static IEnumerable<string> DataRouteIds => OfficeConversionCapabilityCatalog.All
        .Where(route => route.PackageId is "OfficeIMO.Data.Conversions" or "OfficeIMO.Data.Conversions.Excel")
        .Select(route => route.Id);

    private static bool IsDataRoute(string id) => DataRouteIds.Contains(id, StringComparer.OrdinalIgnoreCase);

    private static OfficeConversionReport ConvertData(string id, Stream input, Stream output, OfficeConversionOptions settings, CancellationToken token) {
        int dash = id.IndexOf('-');
        string sourceId = id.Substring(0, dash), targetId = id.Substring(dash + 1);
        DataConversionOptions options = settings.Data?.Clone() ?? new DataConversionOptions();
        options.FailOnLoss |= settings.FailOnLoss;
        if (sourceId == "xlsx") {
            var context = new DataConversionContext(DataConversionEngine.Normalize(options, DataFormat(targetId)!.Value), token);
            using DataEventReader reader = ExcelDataConversions.OpenWorkbook(input, context);
            using DataEventWriter writer = DataConversionEngine.CreateWriter(output, DataFormat(targetId)!.Value, context, 1, Data.Conversions.DataFormat.Csv);
            return DataConversionEngine.Run(reader, writer, context, "Excel", targetId.ToUpperInvariant());
        }

        if (targetId == "xlsx") {
            using StructuredDataReader rows = StructuredDataReader.Create(Seekable(input, out Stream? spooled), DataFormat(sourceId)!.Value, options, owned: spooled, text: false);
            return ExcelDataConversions.WriteRows(rows, output, sourceId.ToUpperInvariant());
        }

        return DataConversionEngine.Convert(input, DataFormat(sourceId)!.Value, output, DataFormat(targetId)!.Value, options, token);
    }

    private static Stream Seekable(Stream input, out Stream? spooled) {
        spooled = null;
        if (input.CanSeek) return input;
        var copy = new MemoryStream();
        input.CopyTo(copy);
        copy.Position = 0;
        spooled = copy;
        return copy;
    }

    private static Data.Conversions.DataFormat? DataFormat(string id) => id switch {
        "json" => Data.Conversions.DataFormat.Json,
        "jsonl" => Data.Conversions.DataFormat.JsonLines,
        "yaml" => Data.Conversions.DataFormat.Yaml,
        "xml" => Data.Conversions.DataFormat.Xml,
        "csv" => Data.Conversions.DataFormat.Csv,
        "tsv" => Data.Conversions.DataFormat.Tsv,
        "toml" => Data.Conversions.DataFormat.Toml,
        _ => null
    };
}

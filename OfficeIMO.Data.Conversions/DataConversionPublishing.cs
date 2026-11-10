using System.Text.Json.Nodes;

namespace OfficeIMO.Data.Conversions;

/// <summary>Shared plumbing of the public <c>To{Target}</c>/<c>SaveAs{Target}</c> methods.</summary>
internal static class DataConversionPublishing {
    public delegate DataEventReader ReaderFactory(DataConversionContext context);

    public static OfficeConversionReport Save(ReaderFactory open, long? documents, DataFormat source, DataFormat target,
        Stream output, DataConversionOptions? options) {
        if (output == null) throw new ArgumentNullException(nameof(output));
        DataConversionOptions effective = DataConversionEngine.Normalize(options, target);
        var context = new DataConversionContext(effective);
        using DataEventReader reader = open(context);
        using DataEventWriter writer = DataConversionEngine.CreateWriter(output, target, context, documents, source);
        return DataConversionEngine.Run(reader, writer, context, DataFormats.Name(source), DataFormats.Name(target));
    }

    public static OfficeConversionReport Save(ReaderFactory open, long? documents, DataFormat source, DataFormat target,
        string path, DataConversionOptions? options) =>
        DataConversionEngine.SaveToPath(path, stream => Save(open, documents, source, target, stream, options));

    /// <summary>Writes the rows of a model as CSV. The model is walked twice: once for the columns, once for the rows.</summary>
    public static OfficeConversionReport SaveCsv(ReaderFactory open, long? documents, DataFormat source, Stream output, DataConversionOptions? options) {
        if (output == null) throw new ArgumentNullException(nameof(output));
        var context = new DataConversionContext(options ?? DataConversionOptions.Default);
        using var rows = new StructuredDataReader(open, documents, context, text: true, owned: null);
        return CsvTableOutput.Write(rows, output, DataFormats.Name(source));
    }

    public static OfficeConversionReport SaveCsv(ReaderFactory open, long? documents, DataFormat source, string path, DataConversionOptions? options) =>
        DataConversionEngine.SaveToPath(path, stream => SaveCsv(open, documents, source, stream, options));

    public static OfficeIMO.CSV.CsvDocument ToCsv(ReaderFactory open, long? documents, DataFormat source, DataConversionOptions? options,
        out OfficeConversionReport report) {
        using var buffer = new MemoryStream();
        report = SaveCsv(open, documents, source, buffer, options);
        buffer.Position = 0;
        return OfficeIMO.CSV.CsvDocument.Load(buffer);
    }

    public static JsonNode? ToJson(ReaderFactory open, long? documents, DataFormat source, DataConversionOptions? options,
        out OfficeConversionReport report) {
        DataConversionOptions effective = options ?? DataConversionOptions.Default;
        if (effective.JsonLines) {
            effective = effective.Clone();
            effective.JsonLines = false;
        }

        using var buffer = new MemoryStream();
        report = Save(open, documents, source, DataFormat.Json, buffer, effective);
        buffer.Position = 0;
        return JsonNode.Parse(buffer);
    }

    public static System.Xml.Linq.XDocument ToXml(ReaderFactory open, long? documents, DataFormat source, DataConversionOptions? options,
        out OfficeConversionReport report) {
        var context = new DataConversionContext(options ?? DataConversionOptions.Default);
        var document = new System.Xml.Linq.XDocument();
        using (System.Xml.XmlWriter xml = document.CreateWriter())
        using (DataEventReader reader = open(context))
        using (var writer = new XmlEventWriter(xml, documents, source == DataFormat.Xml, context, ownsWriter: false)) {
            report = DataConversionEngine.Run(reader, writer, context, DataFormats.Name(source), "XML");
        }

        return document;
    }

    public static OfficeIMO.Yaml.YamlDocument ToYaml(ReaderFactory open, long? documents, DataFormat source, DataConversionOptions? options,
        out OfficeConversionReport report) {
        using var buffer = new MemoryStream();
        report = Save(open, documents, source, DataFormat.Yaml, buffer, options);
        buffer.Position = 0;
        return OfficeIMO.Yaml.YamlDocument.Load(buffer);
    }

    /// <summary>Writes a source to a target file, routing tables through the CSV writer.</summary>
    public static OfficeConversionReport SaveAs(ReaderFactory open, long? documents, DataFormat source, DataFormat target, string path, DataConversionOptions? options) {
        if (target == DataFormat.Json && DataFormats.FromExtension(path) == DataFormat.JsonLines) target = DataFormat.JsonLines;
        return target == DataFormat.Csv ? SaveCsv(open, documents, source, path, options) : Save(open, documents, source, target, path, options);
    }

    public static OfficeConversionReport SaveAs(ReaderFactory open, long? documents, DataFormat source, DataFormat target, Stream stream, DataConversionOptions? options) =>
        target == DataFormat.Csv ? SaveCsv(open, documents, source, stream, options) : Save(open, documents, source, target, stream, options);

    /// <summary>Serializes a JSON node to UTF-8 so it streams through the JSON reader like a file.</summary>
    public static MemoryStream Serialize(JsonNode? node) {
        var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer)) {
            if (node is null) writer.WriteNullValue();
            else node.WriteTo(writer);
        }

        buffer.Position = 0;
        return buffer;
    }
}

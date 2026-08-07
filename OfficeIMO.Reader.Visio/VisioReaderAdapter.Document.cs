using OfficeIMO.Visio;
using System.Text.Json;

namespace OfficeIMO.Reader.Visio;

internal static partial class VisioReaderAdapter {
    /// <summary>
    /// Reads a Visio file and returns the shared OfficeIMO read result envelope.
    /// </summary>
    public static OfficeDocumentReadResult ReadDocument(string visioPath, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        if (visioPath == null) throw new ArgumentNullException(nameof(visioPath));
        if (visioPath.Length == 0) throw new ArgumentException("Visio path cannot be empty.", nameof(visioPath));
        if (!File.Exists(visioPath)) throw new FileNotFoundException($"Visio file '{visioPath}' doesn't exist.", visioPath);

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveVisioOptions = ReaderVisioOptionsCloner.CloneOrDefault(visioOptions);
        ReaderInputLimits.EnforceFileSize(visioPath, effectiveReaderOptions.MaxInputBytes);
        SourceMetadata source = BuildSourceMetadataFromPath(visioPath, effectiveReaderOptions.ComputeHashes);
        VisioDocument document = VisioDocument.Load(visioPath, CreateLoadOptions(effectiveReaderOptions));
        return BuildDocumentResult(document, source, effectiveReaderOptions, effectiveVisioOptions, cancellationToken);
    }

    /// <summary>
    /// Reads a Visio stream and returns the shared OfficeIMO read result envelope.
    /// </summary>
    public static OfficeDocumentReadResult ReadDocument(Stream visioStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        if (visioStream == null) throw new ArgumentNullException(nameof(visioStream));
        if (!visioStream.CanRead) throw new ArgumentException("Visio stream must be readable.", nameof(visioStream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveVisioOptions = ReaderVisioOptionsCloner.CloneOrDefault(visioOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.vsdx");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        Stream parseStream = EnsureParseStream(visioStream, effectiveReaderOptions.MaxInputBytes, cancellationToken, out bool ownsParseStream);
        try {
            UpdateSourceMetadataFromSeekableStream(source, parseStream, effectiveReaderOptions.ComputeHashes);
            if (parseStream.CanSeek) {
                parseStream.Position = 0;
            }

            VisioDocument document = VisioDocument.Load(parseStream, CreateLoadOptions(effectiveReaderOptions));
            return BuildDocumentResult(document, source, effectiveReaderOptions, effectiveVisioOptions, cancellationToken);
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }
    }

    /// <summary>
    /// Converts an already loaded Visio document into the shared OfficeIMO read result envelope.
    /// </summary>
    public static OfficeDocumentReadResult ReadDocument(VisioDocument document, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveVisioOptions = ReaderVisioOptionsCloner.CloneOrDefault(visioOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName ?? document.FilePath, "document.vsdx");
        string sourceIdKey = sourceName == null && !string.IsNullOrWhiteSpace(document.FilePath)
            ? NormalizePathForId(document.FilePath!)
            : logicalSourceName;
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(sourceIdKey)
        };

        return BuildDocumentResult(document, source, effectiveReaderOptions, effectiveVisioOptions, cancellationToken);
    }

    /// <summary>
    /// Reads a Visio file and returns shape-data tables in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTable> ReadTables(string visioPath, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractTables(Read(visioPath, readerOptions, visioOptions, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads a Visio stream and returns shape-data tables in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTable> ReadTables(Stream visioStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractTables(Read(visioStream, sourceName, readerOptions, visioOptions, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Converts an already loaded Visio document into shape-data tables in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTable> ReadTables(VisioDocument document, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractTables(Read(document, sourceName, readerOptions, visioOptions, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads a Visio file and returns shape-data table export payloads in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTableExportBundle> ReadTableExports(string visioPath, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExportTables(ReadTables(visioPath, readerOptions, visioOptions, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Reads a Visio stream and returns shape-data table export payloads in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTableExportBundle> ReadTableExports(Stream visioStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExportTables(ReadTables(visioStream, sourceName, readerOptions, visioOptions, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Converts an already loaded Visio document into shape-data table export payloads in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTableExportBundle> ReadTableExports(VisioDocument document, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExportTables(ReadTables(document, sourceName, readerOptions, visioOptions, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Reads a Visio file and returns the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadDocumentJson(string visioPath, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(visioPath, readerOptions, visioOptions, cancellationToken), indented);
    }

    /// <summary>
    /// Reads a Visio stream and returns the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadDocumentJson(Stream visioStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(visioStream, sourceName, readerOptions, visioOptions, cancellationToken), indented);
    }

    /// <summary>
    /// Converts an already loaded Visio document into the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadDocumentJson(VisioDocument document, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(document, sourceName, readerOptions, visioOptions, cancellationToken), indented);
    }

    private static OfficeDocumentReadResult BuildDocumentResult(VisioDocument document, SourceMetadata source, ReaderOptions readerOptions, ReaderVisioOptions visioOptions, CancellationToken cancellationToken) {
        VisioInspectionSnapshot snapshot = document.CreateInspectionSnapshot();
        ReaderChunk[] chunks = Read(document, source, readerOptions, cancellationToken).ToArray();
        OfficeIMO.OfficeDocumentModel model = document.ToOfficeDocumentModel(
            source.Path,
            new VisioDocumentProjectionOptions {
                MaxTableRows = readerOptions.MaxTableRows,
                IncludeSvgPreviewAssets = visioOptions.IncludeSvgPreviewAssets,
                IncludePngPreviewAssets = visioOptions.IncludePngPreviewAssets,
                SvgOptions = visioOptions.SvgOptions,
                PngOptions = visioOptions.PngOptions
            },
            cancellationToken);
        ReaderVisual[] visuals = BuildDocumentVisuals(snapshot, source).ToArray();
        return BuildReaderResult(model, source, chunks, visuals);
    }

    private static IEnumerable<ReaderVisual> BuildDocumentVisuals(VisioInspectionSnapshot snapshot, SourceMetadata source) {
        for (int pageIndex = 0; pageIndex < snapshot.Pages.Count; pageIndex++) {
            VisioInspectionPageSnapshot page = snapshot.Pages[pageIndex];
            string content = SerializeVisioTopology(page);
            yield return new ReaderVisual {
                Kind = "network",
                Language = "officeimo-visio-topology",
                Content = content,
                PayloadHash = ComputeSha256Hex(content),
                SourceName = page.Name,
                Width = InchesToPoints(page.Width),
                Height = InchesToPoints(page.Height),
                PlacedWidth = InchesToPoints(page.Width),
                PlacedHeight = InchesToPoints(page.Height),
                PlacementCount = 1,
                HasGeometry = true,
                IsAxisAligned = true,
                Location = BuildLocation(source, pageIndex, "diagram", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-topology")
            };
        }
    }

    private static string SerializeVisioTopology(VisioInspectionPageSnapshot page) {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream)) {
            writer.WriteStartObject();
            writer.WritePropertyName("page");
            writer.WriteStartObject();
            writer.WriteNumber("id", page.Id);
            writer.WriteString("name", page.Name);
            writer.WriteNumber("width", page.Width);
            writer.WriteNumber("height", page.Height);
            WriteVisioStringArray(writer, "layers", page.Layers);
            writer.WriteEndObject();

            writer.WritePropertyName("nodes");
            writer.WriteStartArray();
            foreach (VisioInspectionShapeSnapshot shape in page.Shapes) {
                writer.WriteStartObject();
                writer.WriteString("id", shape.Id);
                writer.WriteString("name", shape.Name);
                writer.WriteString("text", shape.Text);
                writer.WriteString("type", shape.Type);
                writer.WriteString("master", shape.MasterNameU);
                writer.WriteString("parentId", shape.ParentId);
                writer.WriteNumber("x", shape.PinX);
                writer.WriteNumber("y", shape.PinY);
                writer.WriteNumber("width", shape.Width);
                writer.WriteNumber("height", shape.Height);
                writer.WriteNumber("angle", shape.Angle);
                WriteVisioStringArray(writer, "layers", shape.Layers);
                WriteVisioShapeData(writer, shape.ShapeData);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();

            writer.WritePropertyName("edges");
            writer.WriteStartArray();
            foreach (VisioInspectionConnectorSnapshot connector in page.Connectors) {
                writer.WriteStartObject();
                writer.WriteString("id", connector.Id);
                writer.WriteString("source", connector.FromId);
                writer.WriteString("target", connector.ToId);
                writer.WriteString("kind", connector.Kind);
                writer.WriteString("label", connector.Label);
                writer.WritePropertyName("waypoints");
                writer.WriteStartArray();
                foreach (VisioInspectionWaypointSnapshot point in connector.Waypoints) {
                    writer.WriteStartObject();
                    writer.WriteNumber("x", point.X);
                    writer.WriteNumber("y", point.Y);
                    writer.WriteEndObject();
                }
                writer.WriteEndArray();
                WriteVisioShapeData(writer, connector.ShapeData);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        return Encoding.UTF8.GetString(stream.ToArray());
    }

    private static void WriteVisioStringArray(Utf8JsonWriter writer, string propertyName, IReadOnlyList<string> values) {
        writer.WritePropertyName(propertyName);
        writer.WriteStartArray();
        foreach (string value in values) writer.WriteStringValue(value);
        writer.WriteEndArray();
    }

    private static void WriteVisioShapeData(Utf8JsonWriter writer, IReadOnlyList<VisioInspectionShapeDataSnapshot> rows) {
        writer.WritePropertyName("data");
        writer.WriteStartArray();
        foreach (VisioInspectionShapeDataSnapshot row in rows) {
            writer.WriteStartObject();
            writer.WriteString("name", row.Name);
            writer.WriteString("label", row.Label);
            writer.WriteString("value", row.Value);
            writer.WriteString("type", row.Type);
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
    }

    private static double InchesToPoints(double value) => value * 72D;
}

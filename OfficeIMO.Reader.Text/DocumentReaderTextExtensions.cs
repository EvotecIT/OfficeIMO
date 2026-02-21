using OfficeIMO.CSV;

namespace OfficeIMO.Reader.Text;

/// <summary>
/// Structured text ingestion helpers for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderTextExtensions {
    /// <summary>
    /// Reads structured text content with CSV/JSON/XML-aware chunking.
    /// Other inputs fallback to <see cref="DocumentReader.Read(string, ReaderOptions?, CancellationToken)"/>.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadStructuredText(string path, ReaderOptions? readerOptions = null, StructuredTextReadOptions? structuredOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var options = Normalize(structuredOptions);
        var ext = GetNormalizedExtension(path);

        if (ext == ".csv" || ext == ".tsv") {
            foreach (var chunk in ReadCsv(path, ext, options, cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        if (ext == ".json") {
            foreach (var chunk in ReadJson(path, options, cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        if (ext == ".xml") {
            foreach (var chunk in ReadXml(path, options, cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        foreach (var chunk in DocumentReader.Read(path, readerOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads structured text content from a stream with CSV/JSON/XML-aware chunking.
    /// Other inputs fallback to <see cref="DocumentReader.Read(Stream, string?, ReaderOptions?, CancellationToken)"/>.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadStructuredText(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, StructuredTextReadOptions? structuredOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var options = Normalize(structuredOptions);
        var ext = GetNormalizedExtension(sourceName);

        if (ext == ".csv" || ext == ".tsv") {
            foreach (var chunk in ReadCsv(stream, sourceName, ext, options, cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        if (ext == ".json") {
            foreach (var chunk in ReadJson(stream, sourceName, options, cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        if (ext == ".xml") {
            foreach (var chunk in ReadXml(stream, sourceName, options, cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        foreach (var chunk in DocumentReader.Read(stream, sourceName, readerOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadCsv(string path, string extension, StructuredTextReadOptions options, CancellationToken cancellationToken) {
        var delimiter = extension == ".tsv" ? '\t' : ',';
        var csv = CsvDocument.Load(path, new CsvLoadOptions {
            Delimiter = delimiter,
            HasHeaderRow = options.CsvHeadersInFirstRow,
            Mode = CsvLoadMode.Stream
        });

        foreach (var chunk in ReadCsvDocument(csv, path, options, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadCsv(Stream stream, string? sourceName, string extension, StructuredTextReadOptions options, CancellationToken cancellationToken) {
        var sourcePath = BuildLogicalSourcePath(sourceName, extension == ".tsv" ? "document.tsv" : "document.csv");
        var delimiter = extension == ".tsv" ? '\t' : ',';
        var text = ReadAllText(stream, cancellationToken);
        var csv = CsvDocument.Parse(text, new CsvLoadOptions {
            Delimiter = delimiter,
            HasHeaderRow = options.CsvHeadersInFirstRow,
            Mode = CsvLoadMode.Stream
        });

        foreach (var chunk in ReadCsvDocument(csv, sourcePath, options, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadCsvDocument(CsvDocument csv, string sourcePath, StructuredTextReadOptions options, CancellationToken cancellationToken) {
        var headers = csv.Header.Count > 0
            ? csv.Header.ToArray()
            : new[] { "Column1" };

        var rows = new List<IReadOnlyList<string>>(capacity: options.CsvChunkRows);
        int chunkIndex = 0;
        int rowIndex = 0;

        foreach (var row in csv.AsEnumerable()) {
            cancellationToken.ThrowIfCancellationRequested();

            var values = new string[Math.Max(headers.Length, row.FieldCount)];
            for (int i = 0; i < values.Length; i++) {
                values[i] = i < row.FieldCount
                    ? Convert.ToString(row[i], CultureInfo.InvariantCulture) ?? string.Empty
                    : string.Empty;
            }

            rows.Add(values);
            rowIndex++;

            if (rows.Count >= options.CsvChunkRows) {
                yield return BuildCsvChunk(sourcePath, headers, rows, chunkIndex, rowIndex - rows.Count, options.IncludeCsvMarkdown);
                rows = new List<IReadOnlyList<string>>(capacity: options.CsvChunkRows);
                chunkIndex++;
            }
        }

        if (rows.Count > 0) {
            yield return BuildCsvChunk(sourcePath, headers, rows, chunkIndex, rowIndex - rows.Count, options.IncludeCsvMarkdown);
        }
    }

    private static IEnumerable<ReaderChunk> ReadJson(string path, StructuredTextReadOptions options, CancellationToken cancellationToken) {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (var chunk in ReadJson(fs, path, options, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadJson(Stream stream, string? sourceName, StructuredTextReadOptions options, CancellationToken cancellationToken) {
        var sourcePath = BuildLogicalSourcePath(sourceName, "document.json");
        JsonDocument? doc = null;
        string? parseError = null;

        try {
            doc = JsonDocument.Parse(stream, new JsonDocumentOptions {
                AllowTrailingCommas = true,
                CommentHandling = JsonCommentHandling.Skip,
                MaxDepth = options.JsonMaxDepth
            });
        } catch (Exception ex) when (ex is not OperationCanceledException) {
            parseError = $"JSON parse error: {ex.GetType().Name}.";
        }

        if (parseError != null) {
            yield return BuildWarningChunk(sourcePath, "json-warning-0000", parseError);
            yield break;
        }

        using (doc) {
            var rows = new List<StructuredRow>(capacity: 1024);
            TraverseJson(doc!.RootElement, "$", depth: 0, maxDepth: options.JsonMaxDepth, rows, cancellationToken);

            foreach (var chunk in BuildStructuredChunks(sourcePath, "json", rows, options.JsonChunkRows, options.IncludeJsonMarkdown, cancellationToken)) {
                yield return chunk;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadXml(string path, StructuredTextReadOptions options, CancellationToken cancellationToken) {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (var chunk in ReadXml(fs, path, options, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadXml(Stream stream, string? sourceName, StructuredTextReadOptions options, CancellationToken cancellationToken) {
        var sourcePath = BuildLogicalSourcePath(sourceName, "document.xml");
        XDocument? doc = null;
        string? parseError = null;

        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Ignore,
                XmlResolver = null
            };

            using var reader = XmlReader.Create(stream, settings);
            doc = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        } catch (Exception ex) when (ex is not OperationCanceledException) {
            parseError = $"XML parse error: {ex.GetType().Name}.";
        }

        if (parseError != null) {
            yield return BuildWarningChunk(sourcePath, "xml-warning-0000", parseError);
            yield break;
        }

        var root = doc!.Root;
        if (root == null) {
            yield return BuildWarningChunk(sourcePath, "xml-warning-0001", "XML document does not contain a root element.");
            yield break;
        }

        var rows = new List<StructuredRow>(capacity: 1024);
        TraverseXml(root, parentPath: string.Empty, rows, cancellationToken);

        foreach (var chunk in BuildStructuredChunks(sourcePath, "xml", rows, options.XmlChunkRows, options.IncludeXmlMarkdown, cancellationToken)) {
            yield return chunk;
        }
    }

    private static string GetNormalizedExtension(string? sourceName) {
        var ext = Path.GetExtension(sourceName ?? string.Empty);
        return ext.ToLowerInvariant();
    }

    private static string BuildLogicalSourcePath(string? sourceName, string defaultName) {
        if (!string.IsNullOrWhiteSpace(sourceName)) return sourceName!;
        return defaultName;
    }

    private static string ReadAllText(Stream stream, CancellationToken cancellationToken) {
        var sb = new StringBuilder();
        var buffer = new char[16 * 1024];
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 16 * 1024, leaveOpen: true);

        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            var read = reader.Read(buffer, 0, buffer.Length);
            if (read <= 0) break;
            sb.Append(buffer, 0, read);
        }

        return sb.ToString();
    }

    private static void TraverseJson(
        JsonElement element,
        string path,
        int depth,
        int maxDepth,
        List<StructuredRow> rows,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();

        if (depth > maxDepth) {
            rows.Add(new StructuredRow(path, "depth-limit", "(max depth reached)"));
            return;
        }

        switch (element.ValueKind) {
            case JsonValueKind.Object:
                if (!element.EnumerateObject().Any()) {
                    rows.Add(new StructuredRow(path, "object", "{}"));
                }

                foreach (var property in element.EnumerateObject()) {
                    var childPath = path == "$" ? "$." + property.Name : path + "." + property.Name;
                    TraverseJson(property.Value, childPath, depth + 1, maxDepth, rows, cancellationToken);
                }

                break;
            case JsonValueKind.Array:
                if (element.GetArrayLength() == 0) {
                    rows.Add(new StructuredRow(path, "array", "[]"));
                }

                int index = 0;
                foreach (var child in element.EnumerateArray()) {
                    var childPath = path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";
                    TraverseJson(child, childPath, depth + 1, maxDepth, rows, cancellationToken);
                    index++;
                }

                break;
            case JsonValueKind.String:
                rows.Add(new StructuredRow(path, "string", element.GetString() ?? string.Empty));
                break;
            case JsonValueKind.Number:
                rows.Add(new StructuredRow(path, "number", element.GetRawText()));
                break;
            case JsonValueKind.True:
            case JsonValueKind.False:
                rows.Add(new StructuredRow(path, "boolean", element.GetRawText()));
                break;
            case JsonValueKind.Null:
                rows.Add(new StructuredRow(path, "null", "null"));
                break;
            default:
                rows.Add(new StructuredRow(path, "value", element.GetRawText()));
                break;
        }
    }

    private static void TraverseXml(XElement element, string parentPath, List<StructuredRow> rows, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();

        var currentPath = BuildXmlPath(element, parentPath);
        var textValue = NormalizeText(GetDirectText(element));
        rows.Add(new StructuredRow(currentPath, "element", textValue));

        foreach (var attribute in element.Attributes()) {
            rows.Add(new StructuredRow(
                currentPath + "/@" + attribute.Name.LocalName,
                "attribute",
                NormalizeText(attribute.Value)));
        }

        foreach (var child in element.Elements()) {
            TraverseXml(child, currentPath, rows, cancellationToken);
        }
    }

    private static IEnumerable<ReaderChunk> BuildStructuredChunks(
        string path,
        string idPrefix,
        IReadOnlyList<StructuredRow> rows,
        int rowsPerChunk,
        bool includeMarkdown,
        CancellationToken cancellationToken) {
        if (rows.Count == 0) {
            yield break;
        }

        int index = 0;
        int chunkIndex = 0;
        while (index < rows.Count) {
            cancellationToken.ThrowIfCancellationRequested();

            int take = Math.Min(rowsPerChunk, rows.Count - index);
            var slice = new List<StructuredRow>(take);
            for (int i = 0; i < take; i++) {
                slice.Add(rows[index + i]);
            }

            var tableRows = slice
                .Select(static r => (IReadOnlyList<string>)new[] { r.Path, r.Type, r.Value })
                .ToArray();

            var table = new ReaderTable {
                Title = Path.GetFileName(path),
                Columns = new[] { "Path", "Type", "Value" },
                Rows = tableRows,
                TotalRowCount = slice.Count,
                Truncated = false
            };

            var plain = BuildPlainStructured(slice);
            var markdown = includeMarkdown ? BuildMarkdownStructured(slice) : null;

            yield return new ReaderChunk {
                Id = string.Concat(idPrefix, "-", chunkIndex.ToString("D4", CultureInfo.InvariantCulture)),
                Kind = ReaderInputKind.Text,
                Location = new ReaderLocation {
                    Path = path,
                    BlockIndex = chunkIndex,
                    SourceBlockIndex = index
                },
                Text = plain,
                Markdown = markdown,
                Tables = new[] { table }
            };

            index += take;
            chunkIndex++;
        }
    }

    private static ReaderChunk BuildCsvChunk(string path, IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows, int chunkIndex, int sourceRowStart, bool includeMarkdown) {
        var table = new ReaderTable {
            Title = Path.GetFileName(path),
            Columns = headers.ToArray(),
            Rows = rows.Select(static r => (IReadOnlyList<string>)r.ToArray()).ToArray(),
            TotalRowCount = rows.Count,
            Truncated = false
        };

        var plain = BuildPlain(headers, rows);
        var markdown = includeMarkdown ? BuildMarkdown(headers, rows) : null;

        return new ReaderChunk {
            Id = string.Concat("csv-", chunkIndex.ToString("D4", CultureInfo.InvariantCulture)),
            Kind = ReaderInputKind.Text,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = chunkIndex,
                SourceBlockIndex = sourceRowStart
            },
            Text = plain,
            Markdown = markdown,
            Tables = new[] { table }
        };
    }

    private static ReaderChunk BuildWarningChunk(string path, string id, string warning) {
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Unknown,
            Location = new ReaderLocation { Path = path },
            Text = warning,
            Warnings = new[] { warning }
        };
    }

    private static string BuildXmlPath(XElement element, string parentPath) {
        var siblingIndex = 1 + element.ElementsBeforeSelf(element.Name).Count();
        var segment = element.Name.LocalName + "[" + siblingIndex.ToString(CultureInfo.InvariantCulture) + "]";
        return parentPath.Length == 0 ? segment : parentPath + "/" + segment;
    }

    private static string GetDirectText(XElement element) {
        var sb = new StringBuilder();
        foreach (var node in element.Nodes().OfType<XText>()) {
            sb.Append(node.Value);
            sb.Append(' ');
        }

        return sb.ToString();
    }

    private static string NormalizeText(string value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        var sb = new StringBuilder(value.Length);
        bool previousWhitespace = false;
        foreach (var ch in value) {
            if (char.IsWhiteSpace(ch)) {
                if (!previousWhitespace) {
                    sb.Append(' ');
                    previousWhitespace = true;
                }
            } else {
                sb.Append(ch);
                previousWhitespace = false;
            }
        }

        var normalized = sb.ToString().Trim();
        if (normalized.Length > 2048) {
            normalized = normalized.Substring(0, 2048);
        }

        return normalized;
    }

    private static string BuildPlain(IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows) {
        var sb = new StringBuilder();
        sb.AppendLine(string.Join(" | ", headers));
        foreach (var row in rows) {
            sb.AppendLine(string.Join(" | ", row));
        }

        return sb.ToString().TrimEnd();
    }

    private static string BuildMarkdown(IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows) {
        var sb = new StringBuilder();

        sb.Append("| ");
        sb.Append(string.Join(" | ", headers.Select(EscapeMarkdownCell)));
        sb.AppendLine(" |");

        sb.Append("| ");
        sb.Append(string.Join(" | ", headers.Select(static _ => "---")));
        sb.AppendLine(" |");

        foreach (var row in rows) {
            sb.Append("| ");
            sb.Append(string.Join(" | ", row.Select(EscapeMarkdownCell)));
            sb.AppendLine(" |");
        }

        return sb.ToString().TrimEnd();
    }

    private static string BuildPlainStructured(IReadOnlyList<StructuredRow> rows) {
        var sb = new StringBuilder();
        sb.AppendLine("Path | Type | Value");
        foreach (var row in rows) {
            sb.Append(row.Path);
            sb.Append(" | ");
            sb.Append(row.Type);
            sb.Append(" | ");
            sb.AppendLine(row.Value);
        }

        return sb.ToString().TrimEnd();
    }

    private static string BuildMarkdownStructured(IReadOnlyList<StructuredRow> rows) {
        var sb = new StringBuilder();
        sb.AppendLine("| Path | Type | Value |");
        sb.AppendLine("| --- | --- | --- |");
        foreach (var row in rows) {
            sb.Append("| ");
            sb.Append(EscapeMarkdownCell(row.Path));
            sb.Append(" | ");
            sb.Append(EscapeMarkdownCell(row.Type));
            sb.Append(" | ");
            sb.Append(EscapeMarkdownCell(row.Value));
            sb.AppendLine(" |");
        }

        return sb.ToString().TrimEnd();
    }

    private static string EscapeMarkdownCell(string value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        return value.Replace("\\", "\\\\").Replace("|", "\\|");
    }

    private static StructuredTextReadOptions Normalize(StructuredTextReadOptions? options) {
        var source = options ?? new StructuredTextReadOptions();
        var normalized = new StructuredTextReadOptions {
            CsvChunkRows = source.CsvChunkRows,
            CsvHeadersInFirstRow = source.CsvHeadersInFirstRow,
            IncludeCsvMarkdown = source.IncludeCsvMarkdown,
            JsonChunkRows = source.JsonChunkRows,
            JsonMaxDepth = source.JsonMaxDepth,
            IncludeJsonMarkdown = source.IncludeJsonMarkdown,
            XmlChunkRows = source.XmlChunkRows,
            IncludeXmlMarkdown = source.IncludeXmlMarkdown
        };

        if (normalized.CsvChunkRows < 1) normalized.CsvChunkRows = 1;
        if (normalized.JsonChunkRows < 1) normalized.JsonChunkRows = 1;
        if (normalized.XmlChunkRows < 1) normalized.XmlChunkRows = 1;
        if (normalized.JsonMaxDepth < 1) normalized.JsonMaxDepth = 1;

        return normalized;
    }

    private sealed class StructuredRow {
        public StructuredRow(string path, string type, string value) {
            Path = path;
            Type = type;
            Value = value;
        }

        public string Path { get; }
        public string Type { get; }
        public string Value { get; }
    }
}

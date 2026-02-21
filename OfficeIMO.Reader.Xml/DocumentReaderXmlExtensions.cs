namespace OfficeIMO.Reader.Xml;

/// <summary>
/// XML ingestion helpers for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderXmlExtensions {
    /// <summary>
    /// Reads XML content from a path with tree-aware chunking.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadXml(string path, ReaderOptions? readerOptions = null, XmlReadOptions? xmlOptions = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(path, effectiveReaderOptions.MaxInputBytes);

        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        foreach (var chunk in ReadXml(fs, path, effectiveReaderOptions, xmlOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads XML content from a stream with tree-aware chunking.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadXml(Stream stream, string? sourceName = null, ReaderOptions? readerOptions = null, XmlReadOptions? xmlOptions = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var options = Normalize(xmlOptions);
        var sourcePath = BuildLogicalSourcePath(sourceName, "document.xml");

        var parseStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            effectiveReaderOptions.MaxInputBytes,
            cancellationToken,
            out var ownsParseStream);

        try {
            XDocument? doc = null;
            string? parseError = null;
            try {
                var settings = new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Ignore,
                    XmlResolver = null
                };

                using var reader = XmlReader.Create(parseStream, settings);
                doc = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            } catch (Exception ex) when (ex is not OperationCanceledException) {
                parseError = "XML parse error: " + ex.GetType().Name + ".";
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

            foreach (var chunk in BuildStructuredChunks(sourcePath, rows, options, cancellationToken)) {
                yield return chunk;
            }
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }
    }

    private static IEnumerable<ReaderChunk> BuildStructuredChunks(
        string path,
        IReadOnlyList<StructuredRow> rows,
        XmlReadOptions options,
        CancellationToken cancellationToken) {
        if (rows.Count == 0) {
            yield break;
        }

        int index = 0;
        int chunkIndex = 0;
        while (index < rows.Count) {
            cancellationToken.ThrowIfCancellationRequested();

            int take = Math.Min(options.ChunkRows, rows.Count - index);
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

            yield return new ReaderChunk {
                Id = "xml-" + chunkIndex.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Xml,
                Location = new ReaderLocation {
                    Path = path,
                    BlockIndex = chunkIndex,
                    SourceBlockIndex = index
                },
                Text = BuildPlain(slice),
                Markdown = options.IncludeMarkdown ? BuildMarkdown(slice) : null,
                Tables = new[] { table }
            };

            index += take;
            chunkIndex++;
        }
    }

    private static void TraverseXml(XElement element, string parentPath, List<StructuredRow> rows, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();

        var currentPath = BuildXmlPath(element, parentPath);
        rows.Add(new StructuredRow(currentPath, "element", NormalizeText(GetDirectText(element))));

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

    private static ReaderChunk BuildWarningChunk(string path, string id, string warning) {
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Xml,
            Location = new ReaderLocation { Path = path },
            Text = warning,
            Warnings = new[] { warning }
        };
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

    private static string BuildPlain(IReadOnlyList<StructuredRow> rows) {
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

    private static string BuildMarkdown(IReadOnlyList<StructuredRow> rows) {
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

    private static XmlReadOptions Normalize(XmlReadOptions? options) {
        var source = options ?? new XmlReadOptions();

        var normalized = new XmlReadOptions {
            ChunkRows = source.ChunkRows,
            IncludeMarkdown = source.IncludeMarkdown
        };

        if (normalized.ChunkRows < 1) normalized.ChunkRows = 1;

        return normalized;
    }

    private static string BuildLogicalSourcePath(string? sourceName, string defaultName) {
        if (!string.IsNullOrWhiteSpace(sourceName)) {
            return sourceName!;
        }

        return defaultName;
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

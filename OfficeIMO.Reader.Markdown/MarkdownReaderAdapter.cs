using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Markdown;

namespace OfficeIMO.Reader.Markdown;

internal static class MarkdownReaderAdapter {
    internal static ReaderMarkdownOptions Clone(ReaderMarkdownOptions? source) => new ReaderMarkdownOptions {
        ChunkByHeadings = source?.ChunkByHeadings ?? true,
        ParserOptions = (source?.ParserOptions ?? new MarkdownReaderOptions()).Clone()
    };

    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions readerOptions, ReaderMarkdownOptions options, CancellationToken cancellationToken) {
        string text = File.ReadAllText(path);
        return Project(text, path, readerOptions, options, cancellationToken);
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? sourceName, ReaderOptions readerOptions, ReaderMarkdownOptions options, CancellationToken cancellationToken) {
        if (stream.CanSeek) stream.Position = 0;
        using var reader = new StreamReader(stream, Encoding.UTF8, true, 4096, leaveOpen: true);
        string text = reader.ReadToEnd();
        return Project(text, string.IsNullOrWhiteSpace(sourceName) ? "document.md" : sourceName!, readerOptions, options, cancellationToken);
    }

    private static OfficeDocumentReadResult Project(string text, string sourceName, ReaderOptions readerOptions, ReaderMarkdownOptions options, CancellationToken cancellationToken) {
        MarkdownReaderOptions parserOptions = (options.ParserOptions ?? new MarkdownReaderOptions()).Clone();
        MarkdownParseResult parsed = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(text ?? string.Empty, parserOptions);
        var projected = new List<ProjectedBlock>();
        var headingStack = new List<HeadingState>();
        var headingSlugs = new Dictionary<string, int>(StringComparer.Ordinal);
        int tableIndex = 0;

        for (int blockIndex = 0; blockIndex < parsed.Document.Blocks.Count; blockIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            IMarkdownBlock block = parsed.Document.Blocks[blockIndex];
            if (block is HeadingBlock heading) {
                UpdateHeadingStack(headingStack, heading.Level, heading.Text, BuildHeadingSlug(heading.Text, headingSlugs));
            }

            string markdown = Normalize(block.RenderMarkdown()).TrimEnd();
            if (markdown.Length == 0) continue;
            MarkdownSyntaxNode? syntax = parsed.FindFinalNodeForAssociatedObject(block);
            MarkdownSourceSpan? span = syntax?.SourceSpan;
            string? headingPath = BuildHeadingPath(headingStack);
            string? hierarchyHeadingPath = ReaderHeadingPath.Combine(headingStack.Select(static item => item.Text));
            string? headingSlug = headingStack.Count == 0 ? null : headingStack[headingStack.Count - 1].Slug;
            string blockKind = BlockKind(block);
            string anchor = BuildAnchor(headingSlug, blockKind, blockIndex, block is HeadingBlock);
            ReaderTable[] tables = block is TableBlock table
                ? new[] { MapTable(table, sourceName, blockIndex, tableIndex, span, headingPath, headingSlug, anchor, readerOptions.MaxTableRows) }
                : block is CodeBlock dataView && string.Equals(dataView.Language, "ix-dataview", StringComparison.OrdinalIgnoreCase) && TryMapIxDataViewTable(dataView.Content, sourceName, blockIndex, tableIndex, span, headingPath, headingSlug, anchor, readerOptions.MaxTableRows, out ReaderTable? projectedTable)
                    ? new[] { projectedTable! }
                : Array.Empty<ReaderTable>();
            if (tables.Length > 0) tableIndex++;
            ReaderVisual[] visuals = block is CodeBlock code && TryMapVisual(code, sourceName, blockIndex, span, headingPath, headingSlug, anchor, out ReaderVisual? visual)
                ? new[] { visual! }
                : Array.Empty<ReaderVisual>();
            projected.Add(new ProjectedBlock(blockIndex, markdown, headingPath, hierarchyHeadingPath, headingSlug, blockKind, anchor, span, block is HeadingBlock, tables, visuals, warnings: null));
        }

        ReaderChunk[] chunks = BuildChunks(projected, sourceName, readerOptions.MaxChars, options.ChunkByHeadings).ToArray();
        return DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.Markdown,
            source: null,
            capabilities: new[] { OfficeDocumentReaderBuilderMarkdownExtensions.HandlerId });
    }

    private static IEnumerable<ReaderChunk> BuildChunks(IReadOnlyList<ProjectedBlock> blocks, string sourceName, int maxChars, bool chunkByHeadings) {
        int limit = Math.Max(256, maxChars);
        int emitted = 0;
        var current = new List<ProjectedBlock>();
        int currentLength = 0;

        foreach (ProjectedBlock block in blocks) {
            bool newSection = chunkByHeadings && block.IsHeading && current.Count > 0;
            bool exceeds = current.Count > 0 && currentLength + 2 + block.Markdown.Length > limit;
            if (newSection || exceeds) {
                yield return CreateChunk(current, sourceName, emitted++);
                current.Clear();
                currentLength = 0;
            }
            if (block.Markdown.Length > limit) {
                if (current.Count > 0) {
                    yield return CreateChunk(current, sourceName, emitted++);
                    current.Clear();
                    currentLength = 0;
                }
                yield return CreateChunk(new[] { block.WithWarning("A single markdown block exceeded MaxChars and was preserved as one chunk.") }, sourceName, emitted++);
                continue;
            }
            current.Add(block);
            currentLength += (current.Count == 1 ? 0 : 2) + block.Markdown.Length;
        }
        if (current.Count > 0) yield return CreateChunk(current, sourceName, emitted);
    }

    private static ReaderChunk CreateChunk(IReadOnlyList<ProjectedBlock> blocks, string sourceName, int emittedIndex) {
        ProjectedBlock first = blocks[0];
        ProjectedBlock last = blocks[blocks.Count - 1];
        string markdown = string.Join("\n\n", blocks.Select(static block => block.Markdown));
        ReaderTable[] tables = blocks.SelectMany(static block => block.Tables).ToArray();
        ReaderVisual[] visuals = blocks.SelectMany(static block => block.Visuals).ToArray();
        return new ReaderChunk {
            Id = $"markdown:{Path.GetFileName(sourceName)}:{emittedIndex.ToString("D4", CultureInfo.InvariantCulture)}",
            Kind = ReaderInputKind.Markdown,
            Location = new ReaderLocation {
                Path = sourceName,
                BlockIndex = emittedIndex,
                SourceBlockIndex = first.Index,
                StartLine = first.Span?.StartLine,
                EndLine = last.Span?.EndLine,
                NormalizedStartLine = first.Span?.StartLine,
                NormalizedEndLine = last.Span?.EndLine,
                HeadingPath = first.HeadingPath,
                HierarchyHeadingPath = first.HierarchyHeadingPath,
                HierarchyHeadingDisplayPath = first.HeadingPath,
                HeadingSlug = first.HeadingSlug,
                SourceBlockKind = first.Kind,
                BlockAnchor = first.Anchor
            },
            Text = markdown,
            Markdown = markdown,
            Tables = tables.Length == 0 ? null : tables,
            Visuals = visuals.Length == 0 ? null : visuals,
            Warnings = blocks.SelectMany(static block => block.Warnings ?? Array.Empty<string>()).Distinct(StringComparer.Ordinal).ToArray() is { Length: > 0 } warnings ? warnings : null
        };
    }

    private static ReaderTable MapTable(TableBlock table, string sourceName, int blockIndex, int tableIndex, MarkdownSourceSpan? span, string? headingPath, string? headingSlug, string anchor, int maxRows) {
        int columnCount = Math.Max(table.Headers.Count, table.Rows.Count == 0 ? 0 : table.Rows.Max(static row => row.Count));
        string[] columns = Enumerable.Range(0, columnCount).Select(index => index < table.Headers.Count && !string.IsNullOrWhiteSpace(table.Headers[index]) ? table.Headers[index] : "Column" + (index + 1).ToString(CultureInfo.InvariantCulture)).ToArray();
        int rowLimit = Math.Max(1, maxRows);
        IReadOnlyList<string>[] rows = table.Rows.Take(rowLimit).Select(row => (IReadOnlyList<string>)Enumerable.Range(0, columnCount).Select(index => index < row.Count ? row[index] : string.Empty).ToArray()).ToArray();
        return new ReaderTable {
            Kind = "markdown-table",
            Columns = columns,
            Rows = rows,
            TotalRowCount = table.Rows.Count,
            Truncated = rows.Length < table.Rows.Count,
            ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, rows),
            Location = new ReaderLocation {
                Path = sourceName,
                SourceBlockIndex = blockIndex,
                StartLine = span?.StartLine,
                EndLine = span?.EndLine,
                NormalizedStartLine = span?.StartLine,
                NormalizedEndLine = span?.EndLine,
                HeadingPath = headingPath,
                HeadingSlug = headingSlug,
                SourceBlockKind = "table",
                BlockAnchor = anchor,
                TableIndex = tableIndex
            }
        };
    }

    private static bool TryMapIxDataViewTable(
        string? rawContent,
        string sourceName,
        int blockIndex,
        int tableIndex,
        MarkdownSourceSpan? span,
        string? headingPath,
        string? headingSlug,
        string anchor,
        int maxRows,
        out ReaderTable? table) {
        table = null;
        string payload = (rawContent ?? string.Empty).TrimEnd('\r', '\n');
        if (payload.Length == 0) return false;

        try {
            using JsonDocument document = JsonDocument.Parse(payload);
            JsonElement root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object) return false;

            var columns = new List<string>();
            var rows = new List<IReadOnlyList<string>>();
            if (root.TryGetProperty("columns", out JsonElement columnsElement) && columnsElement.ValueKind == JsonValueKind.Array) {
                foreach (JsonElement value in columnsElement.EnumerateArray()) columns.Add(ReadJsonScalar(value));
            }

            if (root.TryGetProperty("rows", out JsonElement rowsElement) && rowsElement.ValueKind == JsonValueKind.Array) {
                JsonElement[] sourceRows = rowsElement.EnumerateArray().ToArray();
                if (sourceRows.Length > 0 && sourceRows[0].ValueKind == JsonValueKind.Array) {
                    bool columnsWereProvided = columns.Count > 0;
                    IReadOnlyList<string> first = sourceRows[0].EnumerateArray().Select(ReadJsonScalar).ToArray();
                    if (columns.Count == 0) columns.AddRange(first);
                    for (int index = columnsWereProvided ? 0 : 1; index < sourceRows.Length; index++) {
                        if (sourceRows[index].ValueKind == JsonValueKind.Array) rows.Add(NormalizeRow(sourceRows[index].EnumerateArray().Select(ReadJsonScalar).ToArray(), columns.Count));
                    }
                }
            } else if (root.TryGetProperty("records", out JsonElement recordsElement) && recordsElement.ValueKind == JsonValueKind.Array) {
                JsonElement[] records = recordsElement.EnumerateArray().ToArray();
                if (columns.Count == 0 && records.Length > 0 && records[0].ValueKind == JsonValueKind.Object) {
                    columns.AddRange(records[0].EnumerateObject().Select(static property => property.Name));
                }
                foreach (JsonElement record in records) {
                    if (record.ValueKind == JsonValueKind.Object) {
                        rows.Add(columns.Select(column => record.TryGetProperty(column, out JsonElement value) ? ReadJsonScalar(value) : string.Empty).ToArray());
                    } else if (record.ValueKind == JsonValueKind.Array) {
                        rows.Add(NormalizeRow(record.EnumerateArray().Select(ReadJsonScalar).ToArray(), columns.Count));
                    }
                }
            }

            if (columns.Count == 0) return false;
            string kind = ReadJsonString(root, "kind") ?? "ix-dataview";
            int rowLimit = Math.Max(1, maxRows);
            IReadOnlyList<IReadOnlyList<string>> boundedRows = rows.Take(rowLimit).ToArray();
            table = new ReaderTable {
                Title = ReadJsonString(root, "title") ?? kind,
                Kind = kind,
                CallId = ReadJsonString(root, "call_id"),
                Summary = ReadJsonString(root, "summary"),
                PayloadHash = Hash(payload),
                Columns = columns,
                Rows = boundedRows,
                TotalRowCount = rows.Count,
                Truncated = rows.Count > boundedRows.Count,
                ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, boundedRows),
                Location = new ReaderLocation {
                    Path = sourceName,
                    SourceBlockIndex = blockIndex,
                    StartLine = span?.StartLine,
                    EndLine = span?.EndLine,
                    NormalizedStartLine = span?.StartLine,
                    NormalizedEndLine = span?.EndLine,
                    HeadingPath = headingPath,
                    HeadingSlug = headingSlug,
                    SourceBlockKind = "code",
                    BlockAnchor = anchor,
                    TableIndex = tableIndex
                }
            };
            return true;
        } catch (JsonException) {
            return false;
        }
    }

    private static IReadOnlyList<string> NormalizeRow(IReadOnlyList<string> row, int columnCount) =>
        Enumerable.Range(0, columnCount).Select(index => index < row.Count ? row[index] : string.Empty).ToArray();

    private static string? ReadJsonString(JsonElement root, string propertyName) =>
        root.TryGetProperty(propertyName, out JsonElement value) && value.ValueKind != JsonValueKind.Null
            ? ReadJsonScalar(value)
            : null;

    private static string ReadJsonScalar(JsonElement value) => value.ValueKind switch {
        JsonValueKind.String => value.GetString() ?? string.Empty,
        JsonValueKind.Number => value.GetRawText(),
        JsonValueKind.True => "true",
        JsonValueKind.False => "false",
        JsonValueKind.Null => string.Empty,
        _ => value.GetRawText()
    };

    private static bool TryMapVisual(CodeBlock code, string sourceName, int blockIndex, MarkdownSourceSpan? span, string? headingPath, string? headingSlug, string anchor, out ReaderVisual? visual) {
        string language = (code.Language ?? string.Empty).Trim();
        string? kind = language.ToLowerInvariant() switch {
            "mermaid" => "mermaid",
            "chart" or "ix-chart" => "chart",
            "network" or "ix-network" or "visnetwork" => "network",
            _ => null
        };
        if (kind == null) { visual = null; return false; }
        visual = new ReaderVisual {
            Kind = kind,
            Language = language,
            Content = code.Content ?? string.Empty,
            PayloadHash = Hash(code.Content ?? string.Empty),
            Location = new ReaderLocation {
                Path = sourceName,
                SourceBlockIndex = blockIndex,
                StartLine = span?.StartLine,
                EndLine = span?.EndLine,
                NormalizedStartLine = span?.StartLine,
                NormalizedEndLine = span?.EndLine,
                HeadingPath = headingPath,
                HeadingSlug = headingSlug,
                SourceBlockKind = "code",
                BlockAnchor = anchor
            }
        };
        return true;
    }

    private static string BlockKind(IMarkdownBlock block) => block switch {
        HeadingBlock => "heading",
        TableBlock => "table",
        CodeBlock => "code",
        _ => block.GetType().Name.Replace("Block", string.Empty).ToLowerInvariant()
    };

    private static string? BuildHeadingPath(IReadOnlyList<HeadingState> headings) {
        string[] values = headings.Where(static item => !string.IsNullOrWhiteSpace(item.Text)).Select(static item => item.Text).ToArray();
        return values.Length == 0 ? null : string.Join(" > ", values);
    }

    private static string BuildAnchor(string? headingSlug, string kind, int index, bool isHeading) {
        string normalizedKind = string.IsNullOrWhiteSpace(kind) ? "block" : kind.Trim().ToLowerInvariant();
        if (isHeading && !string.IsNullOrWhiteSpace(headingSlug)) return headingSlug!;
        return !string.IsNullOrWhiteSpace(headingSlug)
            ? headingSlug + "--" + normalizedKind + "-" + index.ToString(CultureInfo.InvariantCulture)
            : normalizedKind + "-" + index.ToString(CultureInfo.InvariantCulture);
    }

    private static void UpdateHeadingStack(List<HeadingState> stack, int level, string text, string slug) {
        for (int index = stack.Count - 1; index >= 0; index--) {
            if (stack[index].Level >= level) stack.RemoveAt(index);
        }
        stack.Add(new HeadingState(level, CollapseWhitespace(text), slug));
    }

    private static string BuildHeadingSlug(string text, IDictionary<string, int> registry) {
        string input = text ?? string.Empty;
        var builder = new StringBuilder(input.Length);
        bool separator = false;
        foreach (char character in input.ToLowerInvariant()) {
            if ((character >= 'a' && character <= 'z') || (character >= '0' && character <= '9')) {
                builder.Append(character);
                separator = false;
            } else if (character == ' ' || character == '-' || character == '_') {
                if (!separator) builder.Append('-');
                separator = true;
            }
        }
        string slug = builder.ToString().Trim('-');
        if (slug.Length == 0) slug = input.Length == 0 ? "heading" : "heading-" + Hash(input).Substring(0, 8);
        if (!registry.TryGetValue(slug, out int count)) {
            registry[slug] = 0;
            return slug;
        }
        int next = count + 1;
        while (registry.ContainsKey(slug + "-" + next.ToString(CultureInfo.InvariantCulture))) next++;
        string candidate = slug + "-" + next.ToString(CultureInfo.InvariantCulture);
        registry[slug] = next;
        registry[candidate] = 0;
        return candidate;
    }

    private static string CollapseWhitespace(string value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        var builder = new StringBuilder(value.Length);
        bool pendingSpace = false;
        foreach (char character in value) {
            if (char.IsWhiteSpace(character)) {
                pendingSpace = builder.Length > 0;
            } else {
                if (pendingSpace) builder.Append(' ');
                builder.Append(character);
                pendingSpace = false;
            }
        }
        return builder.ToString();
    }

    private static string Hash(string value) {
        using SHA256 algorithm = SHA256.Create();
        return string.Concat(algorithm.ComputeHash(Encoding.UTF8.GetBytes(value)).Take(8).Select(static value => value.ToString("x2", CultureInfo.InvariantCulture)));
    }

    private static string Normalize(string value) => (value ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');

    private sealed class HeadingState {
        internal HeadingState(int level, string text, string slug) { Level = level; Text = text; Slug = slug; }
        internal int Level { get; }
        internal string Text { get; }
        internal string Slug { get; }
    }

    private sealed class ProjectedBlock {
        internal ProjectedBlock(int index, string markdown, string? headingPath, string? hierarchyHeadingPath, string? headingSlug, string kind, string anchor, MarkdownSourceSpan? span, bool isHeading, IReadOnlyList<ReaderTable> tables, IReadOnlyList<ReaderVisual> visuals, IReadOnlyList<string>? warnings) {
            Index = index; Markdown = markdown; HeadingPath = headingPath; HierarchyHeadingPath = hierarchyHeadingPath; HeadingSlug = headingSlug; Kind = kind; Anchor = anchor; Span = span; IsHeading = isHeading; Tables = tables; Visuals = visuals; Warnings = warnings;
        }
        internal int Index { get; }
        internal string Markdown { get; }
        internal string? HeadingPath { get; }
        internal string? HierarchyHeadingPath { get; }
        internal string? HeadingSlug { get; }
        internal string Kind { get; }
        internal string Anchor { get; }
        internal MarkdownSourceSpan? Span { get; }
        internal bool IsHeading { get; }
        internal IReadOnlyList<ReaderTable> Tables { get; }
        internal IReadOnlyList<ReaderVisual> Visuals { get; }
        internal IReadOnlyList<string>? Warnings { get; }
        internal ProjectedBlock WithWarning(string warning) => new ProjectedBlock(Index, Markdown, HeadingPath, HierarchyHeadingPath, HeadingSlug, Kind, Anchor, Span, IsHeading, Tables, Visuals, new[] { warning });
    }
}

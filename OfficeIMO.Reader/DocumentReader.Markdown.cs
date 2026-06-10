using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.Reader;

public static partial class DocumentReader {
    private static List<MarkdownChunkBlock> ParseMarkdownBlocksForChunking(string text, ReaderOptions opt, CancellationToken ct) {
        var markdownReaderOptions = CreateMarkdownReaderOptions(opt);
        var parseResult = markdownReaderOptions == null
            ? MarkdownReader.ParseWithSyntaxTree(text ?? string.Empty)
            : MarkdownReader.ParseWithSyntaxTree(text ?? string.Empty, markdownReaderOptions);
        var doc = parseResult.Document;
        var syntaxBlocks = parseResult.SyntaxTree.Children;
        var blocks = new List<MarkdownChunkBlock>(doc.Blocks.Count);
        var headingStack = new List<MarkdownHeadingState>();
        var headingSlugRegistry = new Dictionary<string, int>(StringComparer.Ordinal);
        int nextStartLine = 1;
        int tableIndex = 0;
        bool firstEmittedBlock = true;

        for (int i = 0; i < doc.Blocks.Count; i++) {
            ct.ThrowIfCancellationRequested();

            var block = doc.Blocks[i];
            var syntaxBlock = i < syntaxBlocks.Count ? syntaxBlocks[i] : null;
            var markdown = NormalizeMarkdownLineEndings(block.RenderMarkdown()).TrimEnd();
            if (string.IsNullOrWhiteSpace(markdown)) {
                continue;
            }

            if (!firstEmittedBlock) {
                nextStartLine++;
            }

            int startLine = nextStartLine;
            int endLine = startLine + CountLogicalLines(markdown) - 1;
            int sourceStartLine = syntaxBlock?.SourceSpan?.StartLine ?? startLine;
            int sourceEndLine = syntaxBlock?.SourceSpan?.EndLine ?? endLine;
            bool startsHeading = false;
            if (block is HeadingBlock heading) {
                var slug = BuildMarkdownHeadingSlug(heading.Text, headingSlugRegistry);
                UpdateHeadingStack(headingStack, heading.Level, heading.Text, slug);
                startsHeading = true;
            }

            string? headingPath = BuildHeadingPath(headingStack);
            string? headingSlug = BuildHeadingSlug(headingStack);
            string blockKind = GetMarkdownBlockKind(block);
            string blockAnchor = BuildMarkdownBlockAnchor(headingSlug, blockKind, i, startsHeading);
            IReadOnlyList<ReaderTable> tables = AddMarkdownTableLocations(
                ExtractTables(block, opt),
                ref tableIndex,
                i,
                startLine,
                endLine,
                sourceStartLine,
                sourceEndLine,
                headingPath,
                headingSlug,
                blockKind,
                blockAnchor);
            IReadOnlyList<ReaderVisual> visuals = AddMarkdownVisualLocations(
                ExtractVisuals(block),
                i,
                startLine,
                endLine,
                sourceStartLine,
                sourceEndLine,
                headingPath,
                headingSlug,
                blockKind,
                blockAnchor);

            blocks.Add(new MarkdownChunkBlock(
                blockIndex: i,
                startLine: startLine,
                endLine: endLine,
                sourceStartLine: sourceStartLine,
                sourceEndLine: sourceEndLine,
                headingPath: headingPath,
                headingSlug: headingSlug,
                blockKind: blockKind,
                blockAnchor: blockAnchor,
                markdown: markdown,
                startsHeading: startsHeading,
                tables: tables,
                visuals: visuals));

            nextStartLine += CountLogicalLines(markdown);
            firstEmittedBlock = false;
        }

        return blocks;
    }

    private static MarkdownReaderOptions? CreateMarkdownReaderOptions(ReaderOptions opt) {
        var inputNormalization = CloneMarkdownInputNormalization(opt.MarkdownInputNormalization);
        if (inputNormalization == null) {
            return null;
        }

        return new MarkdownReaderOptions {
            InputNormalization = inputNormalization
        };
    }

    private static IReadOnlyList<ReaderTable> ExtractTables(IMarkdownBlock block, ReaderOptions opt) {
        if (block is TableBlock table) {
            return new[] { MapTable(table, opt) };
        }

        if (block is CodeBlock code &&
            string.Equals(code.Language, "ix-dataview", StringComparison.OrdinalIgnoreCase) &&
            TryMapIxDataViewTable(code.Content, opt, out var dataViewTable) &&
            dataViewTable != null) {
            return new[] { dataViewTable };
        }

        return Array.Empty<ReaderTable>();
    }

    private static IReadOnlyList<ReaderTable> AddMarkdownTableLocations(
        IReadOnlyList<ReaderTable> tables,
        ref int tableIndex,
        int blockIndex,
        int startLine,
        int endLine,
        int sourceStartLine,
        int sourceEndLine,
        string? headingPath,
        string? headingSlug,
        string blockKind,
        string blockAnchor) {
        if (tables == null || tables.Count == 0) {
            return Array.Empty<ReaderTable>();
        }

        var located = new ReaderTable[tables.Count];
        for (int i = 0; i < tables.Count; i++) {
            located[i] = WithMarkdownTableLocation(
                tables[i],
                tableIndex++,
                blockIndex,
                startLine,
                endLine,
                sourceStartLine,
                sourceEndLine,
                headingPath,
                headingSlug,
                blockKind,
                blockAnchor);
        }

        return located;
    }

    private static ReaderTable WithMarkdownTableLocation(
        ReaderTable table,
        int tableIndex,
        int blockIndex,
        int startLine,
        int endLine,
        int sourceStartLine,
        int sourceEndLine,
        string? headingPath,
        string? headingSlug,
        string blockKind,
        string blockAnchor) {
        return new ReaderTable {
            Title = table.Title,
            Kind = table.Kind,
            CallId = table.CallId,
            Summary = table.Summary,
            PayloadHash = table.PayloadHash,
            Location = new ReaderLocation {
                SourceBlockIndex = blockIndex,
                StartLine = sourceStartLine,
                EndLine = sourceEndLine,
                NormalizedStartLine = startLine,
                NormalizedEndLine = endLine,
                HeadingPath = headingPath,
                HeadingSlug = headingSlug,
                SourceBlockKind = string.IsNullOrWhiteSpace(blockKind) ? "table" : blockKind,
                BlockAnchor = blockAnchor,
                TableIndex = tableIndex
            },
            Columns = table.Columns,
            ColumnProfiles = table.ColumnProfiles,
            Rows = table.Rows,
            TotalRowCount = table.TotalRowCount,
            Truncated = table.Truncated
        };
    }

    private static IReadOnlyList<ReaderVisual> ExtractVisuals(IMarkdownBlock block) {
        if (block is CodeBlock code &&
            TryMapVisual(code, out var visual) &&
            visual != null) {
            return new[] { visual };
        }

        return Array.Empty<ReaderVisual>();
    }

    private static IReadOnlyList<ReaderVisual> AddMarkdownVisualLocations(
        IReadOnlyList<ReaderVisual> visuals,
        int blockIndex,
        int startLine,
        int endLine,
        int sourceStartLine,
        int sourceEndLine,
        string? headingPath,
        string? headingSlug,
        string blockKind,
        string blockAnchor) {
        if (visuals == null || visuals.Count == 0) {
            return Array.Empty<ReaderVisual>();
        }

        var located = new ReaderVisual[visuals.Count];
        for (int i = 0; i < visuals.Count; i++) {
            ReaderVisual visual = visuals[i];
            located[i] = new ReaderVisual {
                Kind = visual.Kind,
                Language = visual.Language,
                Content = visual.Content,
                PayloadHash = visual.PayloadHash,
                Location = new ReaderLocation {
                    SourceBlockIndex = blockIndex,
                    StartLine = sourceStartLine,
                    EndLine = sourceEndLine,
                    NormalizedStartLine = startLine,
                    NormalizedEndLine = endLine,
                    HeadingPath = headingPath,
                    HeadingSlug = headingSlug,
                    SourceBlockKind = string.IsNullOrWhiteSpace(blockKind) ? "code" : blockKind,
                    BlockAnchor = blockAnchor
                }
            };
        }

        return located;
    }

    private static bool TryMapVisual(CodeBlock code, out ReaderVisual? visual) {
        visual = null;
        if (code == null) {
            return false;
        }

        var language = (code.Language ?? string.Empty).Trim();
        if (language.Length == 0) {
            return false;
        }

        var normalizedKind = TryNormalizeVisualKind(language);
        if (normalizedKind == null) {
            return false;
        }

        var content = code.Content ?? string.Empty;
        visual = new ReaderVisual {
            Kind = normalizedKind,
            Language = language,
            Content = content,
            PayloadHash = ComputeShortHash(content)
        };
        return true;
    }

    private static string? TryNormalizeVisualKind(string language) {
        if (string.IsNullOrWhiteSpace(language)) {
            return null;
        }

        return language.Trim().ToLowerInvariant() switch {
            "mermaid" => "mermaid",
            "chart" => "chart",
            "ix-chart" => "chart",
            "network" => "network",
            "ix-network" => "network",
            "visnetwork" => "network",
            _ => null
        };
    }

    private static bool TryMapIxDataViewTable(string? rawContent, ReaderOptions opt, out ReaderTable? table) {
        table = null;
        if (string.IsNullOrWhiteSpace(rawContent)) {
            return false;
        }

        try {
            using var document = JsonDocument.Parse(rawContent!);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object) {
                return false;
            }

            if (!TryParseDataViewRows(root, out var parsedColumns, out var parsedRows)) {
                return false;
            }

            var bodyRows = parsedRows.ToList();
            int totalRowCount = bodyRows.Count;

            bool truncated = false;
            if (opt.MaxTableRows > 0 && bodyRows.Count > opt.MaxTableRows) {
                bodyRows = bodyRows.Take(opt.MaxTableRows).ToList();
                truncated = true;
            }

            int columnCount = Math.Max(parsedColumns.Count, bodyRows.Count == 0 ? 0 : bodyRows.Max(static row => row?.Count ?? 0));
            var columns = parsedColumns.Count > 0
                ? EnsureMarkdownTableColumns(parsedColumns, columnCount)
                : BuildMarkdownTableFallbackColumns(columnCount);
            var normalizedRows = bodyRows
                .Select(row => NormalizeMarkdownTableRow(row, columnCount))
                .ToArray();

            table = new ReaderTable {
                Title = TryReadJsonString(root, "title") ?? TryReadJsonString(root, "kind"),
                Kind = TryReadJsonString(root, "kind"),
                CallId = TryReadJsonString(root, "call_id"),
                Summary = TryReadJsonString(root, "summary"),
                PayloadHash = ComputeShortHash(rawContent ?? string.Empty),
                Columns = columns,
                ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, normalizedRows),
                Rows = normalizedRows,
                TotalRowCount = totalRowCount,
                Truncated = truncated
            };
            return true;
        } catch (JsonException) {
            return false;
        }
    }

    private static bool TryParseDataViewRows(JsonElement root, out IReadOnlyList<string> columns, out IReadOnlyList<IReadOnlyList<string>> rows) {
        columns = Array.Empty<string>();
        rows = Array.Empty<IReadOnlyList<string>>();

        if (root.TryGetProperty("rows", out var rowsElement) && rowsElement.ValueKind == JsonValueKind.Array) {
            var parsedRows = new List<IReadOnlyList<string>>();
            foreach (var rowElement in rowsElement.EnumerateArray()) {
                if (rowElement.ValueKind != JsonValueKind.Array) {
                    return false;
                }

                parsedRows.Add(ReadIxDataViewArrayRow(rowElement));
            }

            if (parsedRows.Count == 0) {
                return false;
            }

            columns = parsedRows[0].ToArray();
            rows = parsedRows.Count > 1 ? parsedRows.Skip(1).ToArray() : Array.Empty<IReadOnlyList<string>>();
            return true;
        }

        if (!root.TryGetProperty("records", out var recordsElement) || recordsElement.ValueKind != JsonValueKind.Array) {
            return false;
        }

        var parsedColumns = TryReadIxDataViewColumns(root) ?? DeriveIxDataViewColumnsFromObjectRecords(recordsElement);
        if (parsedColumns == null || parsedColumns.Count == 0) {
            return false;
        }

        var parsedRowsFromRecords = new List<IReadOnlyList<string>>();
        foreach (var recordElement in recordsElement.EnumerateArray()) {
            if (recordElement.ValueKind == JsonValueKind.Array) {
                parsedRowsFromRecords.Add(NormalizeMarkdownTableRow(ReadIxDataViewArrayRow(recordElement), parsedColumns.Count));
                continue;
            }

            if (recordElement.ValueKind == JsonValueKind.Object) {
                parsedRowsFromRecords.Add(ReadIxDataViewObjectRow(recordElement, parsedColumns));
                continue;
            }

            return false;
        }

        columns = parsedColumns;
        rows = parsedRowsFromRecords;
        return true;
    }

    private static IReadOnlyList<string>? TryReadIxDataViewColumns(JsonElement root) {
        if (!root.TryGetProperty("columns", out var columnsElement) || columnsElement.ValueKind != JsonValueKind.Array) {
            return null;
        }

        var columns = new List<string>();
        foreach (var columnElement in columnsElement.EnumerateArray()) {
            columns.Add(ReadIxDataViewScalar(columnElement));
        }

        return columns;
    }

    private static IReadOnlyList<string>? DeriveIxDataViewColumnsFromObjectRecords(JsonElement recordsElement) {
        foreach (var recordElement in recordsElement.EnumerateArray()) {
            if (recordElement.ValueKind != JsonValueKind.Object) {
                continue;
            }

            var columns = new List<string>();
            foreach (var property in recordElement.EnumerateObject()) {
                columns.Add(property.Name);
            }

            return columns.Count == 0 ? null : columns;
        }

        return null;
    }

    private static IReadOnlyList<string> ReadIxDataViewArrayRow(JsonElement rowElement) {
        var row = new List<string>();
        foreach (var cellElement in rowElement.EnumerateArray()) {
            row.Add(ReadIxDataViewScalar(cellElement));
        }

        return row;
    }

    private static IReadOnlyList<string> ReadIxDataViewObjectRow(JsonElement recordElement, IReadOnlyList<string> columns) {
        var row = new string[columns.Count];
        for (int i = 0; i < columns.Count; i++) {
            row[i] = recordElement.TryGetProperty(columns[i], out var cellElement)
                ? ReadIxDataViewScalar(cellElement)
                : string.Empty;
        }

        return row;
    }

    private static string ReadIxDataViewScalar(JsonElement element) {
        return element.ValueKind switch {
            JsonValueKind.String => element.GetString() ?? string.Empty,
            JsonValueKind.Number => element.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Null => string.Empty,
            _ => element.GetRawText()
        };
    }

    private static string? TryReadJsonString(JsonElement root, string propertyName) {
        if (!root.TryGetProperty(propertyName, out var element) || element.ValueKind == JsonValueKind.Null) {
            return null;
        }

        return element.ValueKind == JsonValueKind.String
            ? element.GetString()
            : element.GetRawText();
    }

    private static string ComputeShortHash(string input) {
        var normalized = (input ?? string.Empty).TrimEnd('\r', '\n');
        var data = Encoding.UTF8.GetBytes(normalized);
        byte[] hash;
#if NET8_0_OR_GREATER
        hash = SHA256.HashData(data);
#else
        using (var sha = SHA256.Create()) {
            hash = sha.ComputeHash(data);
        }
#endif

        return ToHex(hash, 8);
    }

    private static string ToHex(byte[] bytes, int take) {
        if (bytes == null || bytes.Length == 0) {
            return string.Empty;
        }

        int len = Math.Min(take, bytes.Length);
        var sb = new StringBuilder(len * 2);
        for (int i = 0; i < len; i++) {
            sb.Append(bytes[i].ToString("x2", CultureInfo.InvariantCulture));
        }

        return sb.ToString();
    }

    private static IReadOnlyList<string> EnsureMarkdownTableColumns(IReadOnlyList<string> headers, int columnCount) {
        if (columnCount <= 0) return Array.Empty<string>();

        var columns = new string[columnCount];
        for (int i = 0; i < columnCount; i++) {
            if (i < headers.Count && !string.IsNullOrWhiteSpace(headers[i])) {
                columns[i] = headers[i];
            } else {
                columns[i] = "Column" + (i + 1).ToString(CultureInfo.InvariantCulture);
            }
        }

        return columns;
    }

    private static IReadOnlyList<string> BuildMarkdownTableFallbackColumns(int columnCount) {
        if (columnCount <= 0) return Array.Empty<string>();

        var columns = new string[columnCount];
        for (int i = 0; i < columnCount; i++) {
            columns[i] = "Column" + (i + 1).ToString(CultureInfo.InvariantCulture);
        }

        return columns;
    }

    private static IReadOnlyList<string> NormalizeMarkdownTableRow(IReadOnlyList<string> row, int columnCount) {
        if (columnCount <= 0) return Array.Empty<string>();

        var values = new string[columnCount];
        for (int i = 0; i < columnCount; i++) {
            values[i] = i < row.Count ? row[i] ?? string.Empty : string.Empty;
        }

        return values;
    }

    private static string NormalizeMarkdownLineEndings(string? markdown) {
        if (string.IsNullOrEmpty(markdown)) return string.Empty;
        return markdown!.Replace("\r\n", "\n").Replace('\r', '\n');
    }

    private static int CountLogicalLines(string markdown) {
        if (string.IsNullOrEmpty(markdown)) return 0;

        int count = 1;
        for (int i = 0; i < markdown.Length; i++) {
            if (markdown[i] == '\n') count++;
        }
        return count;
    }

    private static string GetMarkdownBlockKind(IMarkdownBlock block) {
        if (block == null) return "unknown";

        var name = block.GetType().Name;
        if (name.EndsWith("Block", StringComparison.Ordinal) && name.Length > "Block".Length) {
            name = name.Substring(0, name.Length - "Block".Length);
        }

        return name.ToLowerInvariant();
    }

    private static string BuildMarkdownBlockAnchor(string? headingSlug, string blockKind, int blockIndex, bool startsHeading) {
        var normalizedBlockKind = string.IsNullOrWhiteSpace(blockKind) ? "block" : blockKind.Trim().ToLowerInvariant();
        if (startsHeading && !string.IsNullOrWhiteSpace(headingSlug)) {
            return headingSlug!;
        }

        if (!string.IsNullOrWhiteSpace(headingSlug)) {
            return string.Concat(
                headingSlug!.Trim(),
                "--",
                normalizedBlockKind,
                "-",
                blockIndex.ToString(CultureInfo.InvariantCulture));
        }

        return string.Concat(
            normalizedBlockKind,
            "-",
            blockIndex.ToString(CultureInfo.InvariantCulture));
    }

    private static string BuildMarkdownHeadingSlug(string text, IDictionary<string, int> registry) {
        var input = text ?? string.Empty;
        var sb = new StringBuilder(input.Length);
        bool prevHyphen = false;
        for (int i = 0; i < input.Length; i++) {
            char ch = char.ToLowerInvariant(input[i]);
            if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9')) {
                sb.Append(ch);
                prevHyphen = false;
            } else if (ch == ' ' || ch == '-' || ch == '_') {
                if (!prevHyphen) {
                    sb.Append('-');
                    prevHyphen = true;
                }
            }
        }

        var slug = sb.ToString().Trim('-');
        if (slug.Length == 0) {
            slug = string.IsNullOrEmpty(input)
                ? "heading"
                : "heading-" + ComputeSha256Hex(input).Substring(0, 8);
        }
        if (!registry.TryGetValue(slug, out var count)) {
            registry[slug] = 0;
            return slug;
        }

        int next = count + 1;
        string candidate;
        do {
            candidate = string.IsNullOrEmpty(slug)
                ? "-" + next.ToString(CultureInfo.InvariantCulture)
                : slug + "-" + next.ToString(CultureInfo.InvariantCulture);
            if (!registry.ContainsKey(candidate)) {
                registry[slug] = next;
                registry[candidate] = 0;
                return candidate;
            }
            next++;
        } while (true);
    }

    private static string CollapseWhitespace(string text) {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var sb = new StringBuilder(text.Length);
        bool prevWs = false;
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            bool ws = char.IsWhiteSpace(c);
            if (ws) {
                if (!prevWs) sb.Append(' ');
                prevWs = true;
            } else {
                sb.Append(c);
                prevWs = false;
            }
        }
        return sb.ToString().Trim();
    }

}

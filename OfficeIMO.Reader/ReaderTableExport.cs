using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Reader;

/// <summary>
/// Deterministic text exporters for <see cref="ReaderTable"/> instances.
/// </summary>
public static class ReaderTableExport {
    /// <summary>
    /// Serializes a reader table as RFC 4180-style CSV with CRLF row separators.
    /// </summary>
    /// <param name="table">Table to serialize.</param>
    public static string ToCsv(this ReaderTable table) {
        if (table == null) throw new ArgumentNullException(nameof(table));

        int columnCount = GetColumnCount(table);
        if (columnCount == 0) {
            return string.Empty;
        }

        var builder = new StringBuilder();
        AppendCsvRow(builder, BuildHeaders(table, columnCount));
        IReadOnlyList<IReadOnlyList<string>> rows = table.Rows ?? Array.Empty<IReadOnlyList<string>>();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            builder.Append("\r\n");
            AppendCsvRow(builder, NormalizeRow(rows[rowIndex], columnCount));
        }

        return builder.ToString();
    }

    /// <summary>
    /// Serializes a reader table as a GitHub-style Markdown table.
    /// </summary>
    /// <param name="table">Table to serialize.</param>
    public static string ToMarkdownTable(this ReaderTable table) {
        if (table == null) throw new ArgumentNullException(nameof(table));

        int columnCount = GetColumnCount(table);
        if (columnCount == 0) {
            return string.Empty;
        }

        var builder = new StringBuilder();
        AppendMarkdownRow(builder, BuildHeaders(table, columnCount));
        builder.AppendLine();
        AppendMarkdownSeparator(builder, columnCount);

        IReadOnlyList<IReadOnlyList<string>> rows = table.Rows ?? Array.Empty<IReadOnlyList<string>>();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            builder.AppendLine();
            AppendMarkdownRow(builder, NormalizeRow(rows[rowIndex], columnCount));
        }

        return builder.ToString();
    }

    /// <summary>
    /// Serializes a reader table as deterministic JSON with normalized row width.
    /// </summary>
    /// <param name="table">Table to serialize.</param>
    /// <param name="indented">When true, writes indented JSON for diagnostics and fixtures.</param>
    public static string ToJson(this ReaderTable table, bool indented = false) {
        if (table == null) throw new ArgumentNullException(nameof(table));

        int columnCount = GetColumnCount(table);
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = indented })) {
            writer.WriteStartObject();
            ReaderJsonWriter.WriteNullableString(writer, "title", table.Title);
            ReaderJsonWriter.WriteNullableString(writer, "kind", table.Kind);
            ReaderJsonWriter.WriteNullableString(writer, "callId", table.CallId);
            ReaderJsonWriter.WriteNullableString(writer, "summary", table.Summary);
            ReaderJsonWriter.WriteNullableString(writer, "payloadHash", table.PayloadHash);
            ReaderJsonWriter.WriteLocation(writer, table.Location);
            writer.WriteNumber("totalRowCount", table.TotalRowCount);
            writer.WriteBoolean("truncated", table.Truncated);
            WriteDiagnostics(writer, table.Diagnostics);
            WriteStringArray(writer, "columns", BuildHeaders(table, columnCount));
            WriteRows(writer, table.Rows ?? Array.Empty<IReadOnlyList<string>>(), columnCount);
            WriteColumnProfiles(writer, table.ColumnProfiles ?? Array.Empty<ReaderTableColumnProfile>());
            writer.WriteEndObject();
        }

        return Encoding.UTF8.GetString(stream.ToArray());
    }

    private static void WriteDiagnostics(Utf8JsonWriter writer, ReaderTableDiagnostics? diagnostics) {
        if (diagnostics == null) {
            return;
        }

        writer.WritePropertyName("diagnostics");
        writer.WriteStartObject();
        writer.WriteNumber("confidence", diagnostics.Confidence);
        writer.WriteNumber("schemaConfidence", diagnostics.SchemaConfidence);
        writer.WriteNumber("cellCompleteness", diagnostics.CellCompleteness);
        writer.WriteNumber("columnGeometryConfidence", diagnostics.ColumnGeometryConfidence);
        writer.WriteNumber("sourceRowCount", diagnostics.SourceRowCount);
        writer.WriteNumber("expectedCellCount", diagnostics.ExpectedCellCount);
        writer.WriteNumber("filledCellCount", diagnostics.FilledCellCount);
        writer.WriteNumber("missingCellCount", diagnostics.MissingCellCount);
        writer.WriteNumber("xStart", diagnostics.XStart);
        writer.WriteNumber("xEnd", diagnostics.XEnd);
        writer.WriteNumber("yTop", diagnostics.YTop);
        writer.WriteNumber("yBottom", diagnostics.YBottom);
        writer.WriteNumber("width", diagnostics.Width);
        writer.WriteNumber("height", diagnostics.Height);
        writer.WriteBoolean("hasGeometry", diagnostics.HasGeometry);
        writer.WriteEndObject();
    }

    private static int GetColumnCount(ReaderTable table) {
        int columnCount = table.Columns?.Count ?? 0;
        IReadOnlyList<IReadOnlyList<string>> rows = table.Rows ?? Array.Empty<IReadOnlyList<string>>();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            if (rows[rowIndex] != null && rows[rowIndex].Count > columnCount) {
                columnCount = rows[rowIndex].Count;
            }
        }

        return columnCount;
    }

    private static IReadOnlyList<string> BuildHeaders(ReaderTable table, int columnCount) {
        var headers = new string[columnCount];
        IReadOnlyList<string> columns = table.Columns ?? Array.Empty<string>();
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            string? name = columnIndex < columns.Count ? columns[columnIndex] : null;
            headers[columnIndex] = string.IsNullOrWhiteSpace(name)
                ? "Column" + (columnIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture)
                : name!;
        }

        return headers;
    }

    private static IReadOnlyList<string> NormalizeRow(IReadOnlyList<string>? row, int columnCount) {
        var cells = new string[columnCount];
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            cells[columnIndex] = row != null && columnIndex < row.Count ? row[columnIndex] ?? string.Empty : string.Empty;
        }

        return cells;
    }

    private static void AppendCsvRow(StringBuilder builder, IReadOnlyList<string> cells) {
        for (int i = 0; i < cells.Count; i++) {
            if (i > 0) {
                builder.Append(',');
            }

            AppendCsvCell(builder, cells[i]);
        }
    }

    private static void AppendCsvCell(StringBuilder builder, string? value) {
        string text = value ?? string.Empty;
        bool quote = text.IndexOfAny(new[] { '"', ',', '\r', '\n' }) >= 0;
        if (!quote) {
            builder.Append(text);
            return;
        }

        builder.Append('"');
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            if (c == '"') {
                builder.Append("\"\"");
            } else {
                builder.Append(c);
            }
        }

        builder.Append('"');
    }

    private static void AppendMarkdownRow(StringBuilder builder, IReadOnlyList<string> cells) {
        builder.Append('|');
        for (int i = 0; i < cells.Count; i++) {
            builder.Append(' ');
            builder.Append(EscapeMarkdownCell(cells[i]));
            builder.Append(" |");
        }
    }

    private static void AppendMarkdownSeparator(StringBuilder builder, int columnCount) {
        builder.Append('|');
        for (int i = 0; i < columnCount; i++) {
            builder.Append(" --- |");
        }
    }

    private static void WriteStringArray(Utf8JsonWriter writer, string name, IReadOnlyList<string> values) {
        writer.WritePropertyName(name);
        writer.WriteStartArray();
        for (int i = 0; i < values.Count; i++) {
            writer.WriteStringValue(values[i] ?? string.Empty);
        }

        writer.WriteEndArray();
    }

    private static void WriteRows(Utf8JsonWriter writer, IReadOnlyList<IReadOnlyList<string>> rows, int columnCount) {
        writer.WritePropertyName("rows");
        writer.WriteStartArray();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            IReadOnlyList<string> row = NormalizeRow(rows[rowIndex], columnCount);
            writer.WriteStartArray();
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                writer.WriteStringValue(row[columnIndex]);
            }

            writer.WriteEndArray();
        }

        writer.WriteEndArray();
    }

    private static void WriteColumnProfiles(Utf8JsonWriter writer, IReadOnlyList<ReaderTableColumnProfile> profiles) {
        writer.WritePropertyName("columnProfiles");
        writer.WriteStartArray();
        for (int i = 0; i < profiles.Count; i++) {
            ReaderTableColumnProfile profile = profiles[i];
            writer.WriteStartObject();
            writer.WriteNumber("index", profile.Index);
            writer.WriteString("name", profile.Name ?? string.Empty);
            writer.WriteString("kind", profile.Kind.ToString());
            writer.WriteNumber("nonEmptyCellCount", profile.NonEmptyCellCount);
            writer.WriteNumber("numericCellCount", profile.NumericCellCount);
            writer.WriteNumber("confidence", profile.Confidence);
            writer.WriteEndObject();
        }

        writer.WriteEndArray();
    }

    private static string EscapeMarkdownCell(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        string text = value!;
        var builder = new StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            if (c == '|') {
                builder.Append(@"\|");
            } else if (c == '\r') {
                if (i + 1 < text.Length && text[i + 1] == '\n') {
                    i++;
                }

                builder.Append("<br>");
            } else if (c == '\n') {
                builder.Append("<br>");
            } else {
                builder.Append(c);
            }
        }

        return builder.ToString();
    }
}

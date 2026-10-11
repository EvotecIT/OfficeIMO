using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>Shared table syntax with explicit policies for existing adapter cell contracts.</summary>
internal static class ReaderMarkdownTableWriter {
    internal enum CellStyle { Export, Csv, Literal }

    internal static string Write(IReadOnlyList<string> columns, IEnumerable<IReadOnlyList<string>> rows,
        CellStyle style, CancellationToken cancellationToken = default) {
        var builder = new StringBuilder();
        AppendRow(builder, columns, style, cancellationToken);
        AppendNewLine(builder, style);
        if (columns.Count == 0 && style == CellStyle.Csv) builder.Append("|  |");
        else AppendSeparator(builder, columns.Count, cancellationToken);
        foreach (IReadOnlyList<string> row in rows) {
            AppendNewLine(builder, style);
            AppendRow(builder, row, style, cancellationToken);
        }
        if (style == CellStyle.Literal) builder.Append('\n');
        cancellationToken.ThrowIfCancellationRequested();
        return builder.ToString();
    }

    internal static void AppendRow(StringBuilder builder, IEnumerable<string> cells, CellStyle style,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        builder.Append('|');
        bool any = false;
        foreach (string cell in cells) {
            cancellationToken.ThrowIfCancellationRequested();
            any = true;
            builder.Append(' ');
            AppendCell(builder, cell, style, cancellationToken);
            builder.Append(" |");
        }
        // Preserve the CSV adapter's empty-header representation.
        if (!any && style == CellStyle.Csv) builder.Append("  |");
    }

    internal static void AppendSeparator(StringBuilder builder, int columnCount, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        builder.Append('|');
        for (int index = 0; index < columnCount; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            builder.Append(" --- |");
        }
    }

    private static void AppendNewLine(StringBuilder builder, CellStyle style) {
        if (style == CellStyle.Literal) builder.Append('\n');
        else builder.AppendLine();
    }

    private static void AppendCell(StringBuilder builder, string? value, CellStyle style, CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(value)) return;
        for (int index = 0; index < value!.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            char character = value[index];
            if (style == CellStyle.Literal) {
                if (character == '\r') continue;
                if (character == '\n') { builder.Append(' '); continue; }
                if (character is >= '!' and <= '/' or >= ':' and <= '@' or >= '[' and <= '`' or >= '{' and <= '~') builder.Append('\\');
            } else {
                if (character == '|' || style == CellStyle.Csv && character == '\\') builder.Append('\\');
                if (style == CellStyle.Export && character is '\r' or '\n') {
                    if (character == '\r' && index + 1 < value.Length && value[index + 1] == '\n') index++;
                    builder.Append("<br>");
                    continue;
                }
            }
            builder.Append(character);
        }
        cancellationToken.ThrowIfCancellationRequested();
    }
}
